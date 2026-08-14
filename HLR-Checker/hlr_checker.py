"""
Модуль логики HLR Checker.

Содержит:
- нормализацию телефонных номеров (к формату 8XXXXXXXXXX);
- РЕАЛЬНЫЙ HLR-запрос к SMSC.RU (если задан api_key);
- демо-заглушку (если api_key не задан) — для тестов без списания денег;
- определение оператора и региона (fallback по DEF-коду);
- функцию массовой проверки списка номеров.
"""

import json
import random
import time
import urllib.parse
import urllib.request

from def_codes import OPERATOR_BY_DEF, REGION_BY_DEF


SMSC_BASE_URL = "https://smsc.ru/sys"

# Ключевые статусы HLR в SMSC.RU (см. документацию "Статусы HLR-запроса").
# -3/-2/-1/0  -> запрос ещё не обработан / ошибка запроса
#  1          -> абонент доступен (номер "живой")
#  3          -> номер не существует / заблокирован
#  4          -> абонент временно недоступен
HLR_STATUS_TEXT = {
    1: "активен",
    3: "заблокирован",
    4: "недоступен",
}

HLR_STATUS_RESULT = {
    1: "ok",
    3: "inactive",
    4: "unreachable",
}


def normalize_phone(raw_value):
    """
    Приводит произвольное значение из Excel к нормализованному номеру.

    Поддерживает российские номера:
        +7 900 123-45-67
        8 900 123-45-67
        89001234567
        +79001234567
        9001234567 (10 цифр, добавим 8)

    Возвращает строку в формате "89001234567" или None, если номер не похож на телефон.
    """
    if raw_value is None:
        return None

    # Если это float из Excel (напр. 89001234567.0) — отрежем ".0".
    if isinstance(raw_value, float) and raw_value.is_integer():
        raw_value = str(int(raw_value))
    else:
        raw_value = str(raw_value).strip()

    if not raw_value:
        return None

    # Оставляем только цифры.
    digits = "".join(ch for ch in raw_value if ch.isdigit())

    if not digits:
        return None

    # Приводим к формату 8XXXXXXXXXX (11 цифр, начинается с 8).
    if digits.startswith("8") and len(digits) == 11:
        pass
    elif digits.startswith("7") and len(digits) == 11:
        digits = "8" + digits[1:]
    elif len(digits) == 10:
        digits = "8" + digits
    else:
        return None

    # Итоговая проверка: 11 цифр и начинается с 8.
    if len(digits) != 11 or not digits.startswith("8"):
        return None

    return digits


def get_operator(phone):
    """Определяет оператора по DEF-коду номера (fallback, до-MNP)."""
    def_code = phone[1:4]
    return OPERATOR_BY_DEF.get(def_code, "не определён")


def get_region(phone):
    """Определяет регион регистрации SIM по DEF-коду номера (fallback, до-MNP)."""
    def_code = phone[1:4]
    return REGION_BY_DEF.get(def_code, "не определён")


# ---------------------------------------------------------------------------
# Демо-режим (заглушка)
# ---------------------------------------------------------------------------

def hlr_check_stub(phone):
    """
    ЗАГЛУШКА HLR-запроса.

    Имитирует ответ API без реального обращения к сервису и без списания денег.
    Используется, когда в config.json не задан api_key.

    Для наглядности результат детерминирован по номеру:
    один и тот же номер всегда даёт одинаковый статус в демо-режиме.
    """
    time.sleep(0.03)

    seed = sum(int(ch) for ch in phone if ch.isdigit())
    rnd = random.Random(seed)

    roll = rnd.random()
    if roll < 0.82:
        status = "активен"
        result = "ok"
    elif roll < 0.95:
        status = "заблокирован"
        result = "inactive"
    else:
        status = "недоступен"
        result = "unreachable"

    return {
        "phone": phone,
        "status": status,
        "result": result,
        "operator": get_operator(phone),
        "region": get_region(phone),
    }


# ---------------------------------------------------------------------------
# Реальный режим (SMSC.RU)
# ---------------------------------------------------------------------------

def _http_get_json(url, timeout=30):
    """GET-запрос и разбор JSON-ответа. При ошибке возвращает dict с 'error'."""
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            data = resp.read().decode("utf-8", "ignore")
            return json.loads(data)
    except Exception as e:
        return {"error": str(e), "error_code": None}


def _to_smsc_phone(phone):
    """
    SMSC.RU принимает российские мобильные в формате 7XXXXXXXXXX.
    Наш внутренний формат — 8XXXXXXXXXX, поэтому меняем ведущую 8 на 7.
    """
    if phone.startswith("8"):
        return "7" + phone[1:]
    return phone


def _api_error_result(phone, data):
    """Формирует результат при ошибке API."""
    error = data.get("error") or "неизвестная ошибка API"
    return {
        "phone": phone,
        "status": "ошибка API",
        "result": "api_error",
        "operator": "—",
        "region": "—",
        "error": error,
    }


def _clean_hlr_field(value):
    """
    Очищает текстовое поле HLR от пробелов и мусорных символов.

    SMSC.RU иногда возвращает поля вида " ", ".   " и т.п., которые
    фактически означают «данных нет». Такие значения приводим к пустой строке.
    """
    if value is None:
        return ""
    s = str(value).strip()
    if not s or all(ch in ". -–—," for ch in s):
        return ""
    return s


def _parse_hlr_status(phone, data):
    """Преобразует ответ status.php (all=2) в единый словарь результата."""
    if not data:
        return {
            "phone": phone,
            "status": "не определён",
            "result": "unknown",
            "operator": "—",
            "region": "—",
        }

    status = data.get("status")
    try:
        status = int(status)
    except (TypeError, ValueError):
        status = None

    # SMSC.RU в HLR-ответе отдаёт оператора в поле "net" (а не "operator").
    # Поле "region" часто пустое/мусорное — тогда берём регион по DEF-коду.
    operator = (
        _clean_hlr_field(data.get("net"))
        or _clean_hlr_field(data.get("operator"))
        or get_operator(phone)
    )
    region = _clean_hlr_field(data.get("region"))
    if not region:
        region = get_region(phone)

    text = HLR_STATUS_TEXT.get(status)
    result = HLR_STATUS_RESULT.get(status)

    if text is None:
        # Нестандартный/промежуточный статус — берём текстовое описание API.
        text = data.get("status_name") or data.get("comment") or "не определён"
        result = "unknown"

    return {
        "phone": phone,
        "status": text,
        "result": result,
        "operator": operator,
        "region": region,
        "raw": data,
    }


def hlr_check_real(phone, api_key, timeout=30, poll_interval=2, max_attempts=30):
    """
    Реальный HLR-запрос к SMSC.RU.

    Алгоритм:
        1) send.php?apikey=...&phones=<phone>&hlr=1&fmt=3 -> получаем id.
        2) status.php?apikey=...&phone=<phone>&id=<id>&all=2&fmt=3 -> статус.
        3) Опрашиваем статус, пока он не станет окончательным
           (или не истечёт лимит попыток).

    Возвращает dict того же формата, что и hlr_check_stub().
    """
    # 1. Отправка HLR-запроса.
    send_params = {
        "apikey": api_key,
        "phones": _to_smsc_phone(phone),
        "hlr": 1,
        "fmt": 3,
    }
    send_url = f"{SMSC_BASE_URL}/send.php?" + urllib.parse.urlencode(send_params)
    send_data = _http_get_json(send_url, timeout=timeout)

    if "error" in send_data and send_data.get("id") is None:
        return _api_error_result(phone, send_data)

    hlr_id = send_data.get("id")
    if not hlr_id:
        return _api_error_result(phone, {"error": "в ответе нет id запроса", "error_code": None})

    # 2. Опрос статуса HLR.
    last_data = {}
    for _ in range(max_attempts):
        time.sleep(poll_interval)

        status_params = {
            "apikey": api_key,
            "phone": _to_smsc_phone(phone),
            "id": hlr_id,
            "all": 2,
            "fmt": 3,
        }
        status_url = f"{SMSC_BASE_URL}/status.php?" + urllib.parse.urlencode(status_params)
        status_data = _http_get_json(status_url, timeout=timeout)
        last_data = status_data

        status = status_data.get("status")
        try:
            status = int(status)
        except (TypeError, ValueError):
            status = None

        # 0 — в очереди, -1 — ещё не обработан: ждём дальше.
        if status in (0, -1):
            continue

        # Любой другой статус считаем окончательным (включая 1, 3, 4, ошибки).
        break

    return _parse_hlr_status(phone, last_data)


# ---------------------------------------------------------------------------
# Массовая проверка
# ---------------------------------------------------------------------------

def check_numbers(numbers, api_key=""):
    """
    Принимает список номеров, возвращает список словарей со статусами.

    numbers: список сырых значений (или уже нормализованных строк).
    api_key: если задан — используется реальный HLR (SMSC.RU),
             иначе — демо-заглушка.

    Возвращает список dict вида:
        {"phone": ..., "status": ..., "result": ...,
         "operator": ..., "region": ...}
    """
    results = []
    for raw in numbers:
        phone = normalize_phone(raw)
        if phone is None:
            results.append({
                "phone": raw,
                "status": "ошибка формата",
                "result": "invalid",
                "operator": "—",
                "region": "—",
            })
            continue

        if api_key:
            results.append(hlr_check_real(phone, api_key))
        else:
            results.append(hlr_check_stub(phone))

    return results


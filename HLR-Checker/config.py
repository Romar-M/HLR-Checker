"""
Загрузка конфигурации HLR Checker.

API-ключ хранится в отдельном файле config.json в той же папке,
чтобы секреты не попадали в код и их было легко менять без правки .py.
"""

import json
import os

CONFIG_FILENAME = "config.json"


def _config_path():
    """Возвращает абсолютный путь к config.json рядом с этим модулем."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, CONFIG_FILENAME)


def load_config():
    """Читает config.json и возвращает словарь с настройками."""
    path = _config_path()
    if not os.path.exists(path):
        return {}

    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def get_api_key():
    """Возвращает API-ключ из config.json (пустая строка, если не найден)."""
    return load_config().get("api_key", "")


def mask_api_key(api_key):
    """
    Возвращает замаскированный ключ для безопасного вывода в консоль.

    Пример: 5agi***F989
    """
    if not api_key:
        return ""

    if len(api_key) <= 8:
        return "*" * len(api_key)

    return api_key[:4] + "***" + api_key[-4:]


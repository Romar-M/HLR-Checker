"""
HLR Checker — простая рабочая версия (запуск двойным кликом).

Способы запуска:
    1) Двойной клик по start.bat  (рекомендуется)
    2) python main.py
    3) python main.py path/to/file.xlsx

Логика:
    - API-ключ загружается из отдельного файла config.json.
    - Входные файлы берутся из папки export/ (все .xlsx, кроме временных ~$...).
    - Обработанный файл сохраняется в папку import/ с тем же именем.
    - Пустые столбцы в обработанном файле удаляются.
    - Если задан api_key — номера проверяются реальным HLR (SMSC.RU).
      Если ключ не задан — используется демо-заглушка (без списания денег).
    - Если передан путь через аргумент — обрабатывается именно этот файл,
      а результат также кладётся в import/.
"""

import os
import sys
import time

from config import get_api_key, mask_api_key
from excel_utils import process_workbook


EXPORT_DIR_NAME = "export"
IMPORT_DIR_NAME = "import"


def main():
    print("=" * 60)
    print("[PHONE]  HLR Checker — проверка «живости» номеров")
    print("=" * 60)

    # API-ключ хранится в отдельном файле config.json.
    api_key = get_api_key()
    if api_key:
        print(f"[KEY]   API-ключ загружен из config.json: {mask_api_key(api_key)}")
        print("[MODE]  Реальный режим (SMSC.RU). HLR-запросы платные и занимают время.")
    else:
        print("[MODE]  Демо-режим (заглушка). Реальные HLR-запросы не выполняются.")
        print('[WARN]  API-ключ не найден. Создайте config.json с полем "api_key".')

    script_dir = os.path.dirname(os.path.abspath(__file__))
    export_dir = os.path.join(script_dir, EXPORT_DIR_NAME)
    import_dir = os.path.join(script_dir, IMPORT_DIR_NAME)

    # Папку import создаём заранее, чтобы результат всегда было куда положить.
    os.makedirs(import_dir, exist_ok=True)

    # Определяем список входных файлов.
    if len(sys.argv) > 1:
        input_paths = [sys.argv[1]]
    else:
        input_paths = _find_export_files(export_dir)

    if not input_paths:
        print("\n[ERROR]  Не найдено ни одного .xlsx-файла для обработки.")
        print(f"[HINT]   Положите файл в папку: {export_dir}")
        _pause()
        return

    print(f"\n[IN]    Папка с исходными файлами: {export_dir}")
    print(f"[OUT]   Папка для результатов: {import_dir}")
    print(f"[COUNT] Найдено файлов: {len(input_paths)}\n")

    total_start = time.time()
    ok_count = 0

    for i, input_path in enumerate(input_paths, 1):
        input_path = os.path.abspath(input_path)

        if not os.path.exists(input_path):
            print(f"[{i}/{len(input_paths)}] [ERROR]  Файл не найден: {input_path}")
            continue

        if not input_path.lower().endswith(".xlsx"):
            print(f"[{i}/{len(input_paths)}] [WARN]   Пропущен (не .xlsx): {input_path}")
            continue

        # Результат кладём в import/ с тем же именем файла.
        output_path = os.path.join(import_dir, os.path.basename(input_path))

        print(f"[{i}/{len(input_paths)}] Обработка: {os.path.basename(input_path)}")
        print("      " + "-" * 48)

        start = time.time()
        try:
            removed_cols = process_workbook(input_path, output_path, api_key=api_key)
        except Exception as e:
            print(f"      [ERROR]  Ошибка при обработке: {e}\n")
            continue

        elapsed = time.time() - start
        ok_count += 1
        print(f"      [OK]    Сохранено: {output_path}")
        if removed_cols:
            print(f"      [COLS]  Удалено пустых столбцов: {removed_cols}")
        print(f"      [TIME]  Время обработки: {elapsed:.2f} сек\n")

    total_elapsed = time.time() - total_start

    print("=" * 60)
    print(f"[DONE]  Обработано файлов: {ok_count} из {len(input_paths)}")
    print(f"[TIME]  Общее время: {total_elapsed:.2f} сек")
    print("=" * 60)
    print()

    _pause()


def _find_export_files(export_dir):
    """
    Возвращает отсортированный список .xlsx-файлов из папки export/.

    Игнорируем:
        - временные файлы Excel (~$...);
        - всё, что не заканчивается на .xlsx (в т.ч. desktop.ini).
    """
    if not os.path.isdir(export_dir):
        return []

    files = []
    for name in os.listdir(export_dir):
        # Пропускаем временные файлы, которые Excel создаёт при открытой книге.
        if name.startswith("~$"):
            continue
        if not name.lower().endswith(".xlsx"):
            continue
        files.append(os.path.join(export_dir, name))

    files.sort()
    return files


def _pause():
    """Ожидание нажатия Enter, чтобы окно не закрылось при двойном клике."""
    try:
        input("Нажмите Enter для выхода...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()


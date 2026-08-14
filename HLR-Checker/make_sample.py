"""
Создаёт тестовый Excel-файл для проверки HLR Checker.

Запуск:
    python make_sample.py

Появится файл sample_input.xlsx, который можно скормить main.py.
"""

from openpyxl import Workbook


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Контакты"

    headers = ["Имя", "Рабочий телефон", "Домашний телефон", "Email"]
    rows = [
        ["Иван", "+7 900 123-45-67", "8 495 000-11-22", "ivan@example.com"],
        ["Мария", "89001234568", "+7 900 111-22-33", "maria@example.com"],
        ["Пётр", "не указан", "9001234567", "petr@example.com"],
        ["Ольга", "+7 911 222-33-44", "8 812 555-44-33", "olga@example.com"],
        ["Сергей", "89031112233", "+7 495 777-88-99", "sergey@example.com"],
        ["Анна", "+7 921 000-00-00", "89112223344", "anna@example.com"],
    ]

    ws.append(headers)
    for row in rows:
        ws.append(row)

    wb.save("sample_input.xlsx")
    print("✅ Создан файл sample_input.xlsx")
    print("Запусти проверку: python main.py sample_input.xlsx")


if __name__ == "__main__":
    main()


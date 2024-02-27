import openpyxl
import json
import urllib.parse
import re

def extract_query_from_text(text):
    start_index = text.find("searchQuery=")
    if start_index != -1:
        start_index += len("searchQuery=")
        end_index = text.find('"', start_index)
        if end_index != -1:
            return text[start_index:end_index]
    return None

def decode_to_cyrillic(text):
    return urllib.parse.unquote(text)

def process_excel_file(input_file, output_file):
    data = []

    # Открываем эксель файл
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active

    # Проходимся по всем ячейкам в первом столбце и извлекаем запросы
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1):
        for cell in row:
            query = extract_query_from_text(str(cell.value))
            if query is not None:
                query = decode_to_cyrillic(query)  # Декодируем кириллицу
                # Заменяем "+" на пробелы
                query = query.replace("+", " ")
                # Проверяем, содержит ли строка только числа
                if query.isdigit():
                    continue
                # Убираем последний символ, если он "+"
                if query.endswith(" "):
                    query = query[:-1]
                data.append(query)

    # Записываем данные в JSON файл
    with open(output_file, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False)

# Пример использования функции
input_file = 'input.xlsx'
output_file = 'output.json'

process_excel_file(input_file, output_file)

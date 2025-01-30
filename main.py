import os
import pandas as pd
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook

def fix_xml_file(input_path, output_path):
    """Исправляет XML-файл и сохраняет его в новый файл."""
    try:
        with open(input_path, "rb") as file:
            content = file.read()

        # Попытка автоопределения кодировки
        encodings = ["utf-8", "windows-1251", "ISO-8859-1"]
        for enc in encodings:
            try:
                content = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue

        # Используем html.parser для исправления XML
        soup = BeautifulSoup(content, "html.parser")

        with open(output_path, "w", encoding="utf-8") as file:
            file.write(soup.prettify())

        return True
    except Exception as e:
        print(f"❌ Ошибка при исправлении {input_path}: {e}")
        return False

def extract_date(root):
    """Извлекает дату документа из XML."""
    date_document = "N/A"
    
    # Пытаемся найти дату в элементе <date> в разделе <zglv>
    zglv_date = root.find(".//zglv/date")
    if zglv_date is not None and zglv_date.text:
        date_document = zglv_date.text.strip()
    
    return date_document

def parse_xml_file(file_path):
    """Парсит XML-файл и извлекает данные."""
    try:
        parser = etree.XMLParser(recover=False)  # Отключаем автоматическое исправление
        tree = etree.parse(file_path, parser)
        root = tree.getroot()
    except etree.XMLSyntaxError as e:
        print(f"❌ Ошибка в файле {file_path}: {e}")
        return None, None  # Возвращаем None в случае ошибки

    parsed_data = []

    # Извлекаем данные из элемента <zglv>
    zglv = root.find("zglv")
    version = zglv.find("version").text.strip() if zglv is not None and zglv.find("version") is not None else "N/A"

    # Получаем дату документа
    date_document = extract_date(root)

    # Определяем, какие теги есть в <zap>
    example_zap = root.find("zap")
    if example_zap is None:
        return None, None

    column_names = ["Версия", "Дата документа"]
    first_entry = ["N/A", "N/A"]

    for child in example_zap:
        column_names.append(child.tag)
        first_entry.append(child.text.strip() if child.text else "N/A")

    # Собираем данные
    for zap in root.findall("zap"):
        row = [version, date_document]  # Здесь используем дату для каждого <zap>
        for col in column_names[2:]:  # Пропускаем первые два (они уже добавлены)
            row.append(zap.find(col).text.strip() if zap.find(col) is not None and zap.find(col).text else "N/A")
        parsed_data.append(row)

    return column_names, parsed_data

def process_xml_files(input_folder, output_file, fixed_folder):
    """Обрабатывает XML-файлы: сначала исправляет, затем парсит и записывает в Excel."""
    if not os.path.exists(fixed_folder):
        os.makedirs(fixed_folder)  # Создаем папку для исправленных файлов

    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"

    file_count = 0  # Подсчет обработанных файлов

    for file_name in os.listdir(input_folder):
        file_path = os.path.join(input_folder, file_name)
        fixed_path = os.path.join(fixed_folder, file_name)

        if os.path.isfile(file_path) and file_name.endswith(".xml"):
            print(f"📄 Обрабатываю файл: {file_name}")

            # Исправляем XML перед парсингом
            if fix_xml_file(file_path, fixed_path):
                sheet_name = file_name.replace(".xml", "")[:31]  # Название листа (максимум 31 символ)
                ws = wb.create_sheet(title=sheet_name)

                # Парсим уже исправленный файл
                result = parse_xml_file(fixed_path)
                if not result or result[1] is None:
                    print(f"🚨 Пропущен файл {file_name} из-за ошибки в парсинге.")
                    continue

                headers, data = result

                # Добавляем заголовки
                ws.append(headers)

                # Добавляем данные
                for row in data:
                    ws.append(row)

                file_count += 1  # Увеличиваем счетчик обработанных файлов

    # Проверяем, были ли обработаны файлы
    if file_count == 0:
        print("❌ Ни один XML-файл не был обработан. Проверьте структуру входных данных.")

    wb.save(output_file)
    print(f"✅ Файл успешно сохранен: {output_file}")

# Папки и файлы
input_folder = "/Users/andru_shaa/Desktop/code_spaces/xml_parse/files"  # Оригинальные XML-файлы
fixed_folder = "/Users/andru_shaa/Desktop/code_spaces/xml_parse/fixed_files"  # Исправленные файлы
output_file = "output.xlsx"  # Итоговый Excel-файл

# Запускаем обработку
process_xml_files(input_folder, output_file, fixed_folder)
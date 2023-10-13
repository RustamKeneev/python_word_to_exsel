from docx import Document
import openpyxl
import re
from openpyxl.styles import Alignment

# Открываем документ DOCX
doc = Document("your_document.docx")

# Создаем новый файл Excel
workbook = openpyxl.Workbook()
worksheet = workbook.active

# Индекс строки в Excel
row_index = 2

# Устанавливаем значения по умолчанию
patient_name = ""
patient_birth_year = 0
examination_date = ""

# Заголовки для Excel
column_headers = {
    'A1': 'Пациент',
    'B1': 'Год рождения',
    'C1': 'Дата исследования'
}

# Записываем заголовки в ячейки
for col, header in column_headers.items():
    worksheet[col] = header

for paragraph in doc.paragraphs:
    text = paragraph.text
    if "Пациент: " in text:
        patient_name = text.replace("Пациент: ", "").strip()
    if "Год рождения:" in text:
        clean_text = text[-4:]  # Извлекаем последние четыре символа, предполагая, что это год рождения
        try:
            patient_birth_year = int(clean_text)
        except ValueError:
            print(f"Ошибка: невозможно преобразовать '{clean_text}' к целому числу года рождения.")
            patient_birth_year = 0  # Или другое значение по умолчанию
    elif "Дата исследования:" in text:
        examination_date = text.replace("Дата исследования:", "").strip()

        # Добавляем данные в Excel
        worksheet.cell(row=row_index, column=1, value=patient_name)
        worksheet.cell(row=row_index, column=2, value=patient_birth_year)
        worksheet.cell(row=row_index, column=3, value=examination_date)
        row_index += 1

# Сохраняем файл Excel
workbook.save("output.xlsx")



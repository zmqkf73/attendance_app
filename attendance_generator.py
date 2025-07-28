# attendance_generator.py

import re
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.cell.cell import Cell

def convert_non_string_to_string(x):
    return str(x) if not isinstance(x, str) else x

def remove_newlines(text):
    return text.replace("\n", " ")

def strip_edges(text):
    return text.strip()

def reduce_whitespace(text):
    return re.sub(r"\s+", " ", text)

def remove_all_whitespace(text):
    return re.sub(r"\s+", "", text)

def insert_space_before_brackets(text):
    return re.sub(r"([\w\uac00-\ud7a3])(?=[\(\[\{])", r"\1 ", text)

def replace_tilde_with_dash(text):
    return text.replace("~", "-")

def insert_space_between_adjacent_brackets(text):
    return re.sub(r"(\))(?=\()", r"\1 ", text)

def format_text(text):
    text = convert_non_string_to_string(text)
    text = remove_newlines(text)
    text = insert_space_before_brackets(text)
    text = insert_space_between_adjacent_brackets(text)
    text = replace_tilde_with_dash(text)
    text = reduce_whitespace(text)
    text = strip_edges(text)
    return text

def clean_name(name):
    name = convert_non_string_to_string(name)
    name = remove_newlines(name)
    name = strip_edges(name)
    name = remove_all_whitespace(name)
    return name

def capitalize_first_word_if_english(text):
    match = re.match(r"^([A-Za-z]+)(\s|$)", text)
    if match:
        first = match.group(1)
        rest = text[len(first):]
        return first.capitalize() + rest
    return text

def generate_attendance(records, template_path, year=None, month=None):
    TEMPLATE_ROWS = 31
    TEMPLATE_COLS = 29
    used_block_count = {}

    wb = load_workbook(template_path)
    template_ws = wb["ABC"]

    for record in records:
        teacher = record["강사"]
        if not isinstance(teacher, str) or teacher.strip().lower() == "nan" or not teacher.strip() or teacher == "강사":
            continue

        teacher = teacher.strip()
        course = record["과정"]
        course_class_name = course.split("/")[0] if isinstance(course, str) and course.strip() else ""
        day = record["요일"]
        time = record["시간"]
        students = record["학생목록"]

        if teacher in wb.sheetnames:
            ws = wb[teacher]
        else:
            ws = wb.copy_worksheet(template_ws)
            ws.title = teacher
            used_block_count[teacher] = 0

        block_index = used_block_count.get(teacher, 0)
        start_row = block_index * TEMPLATE_ROWS + 1
        used_block_count[teacher] = block_index + 1

        today = datetime.today()
        year_str = str(year)[2:] if year is not None else today.strftime("%y")
        month_str = str(month) if month is not None else str(today.month)
        formatted_date = f"{year_str}년 {month_str}월"

        ws.cell(row=start_row + 2, column=2).value = formatted_date
        ws.cell(row=start_row + 2, column=7).value = f"{day} {time}"

        for row in ws.iter_rows(min_row=start_row, max_row=start_row + TEMPLATE_ROWS - 1, max_col=TEMPLATE_COLS):
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                if "CLASS:" in cell.value:
                    cell.value = cell.value.split("CLASS:")[0] + f"CLASS: {course_class_name}"
                if "담임" in cell.value and "강사" in cell.value:
                    cell.value = f"담임 강사: {teacher}"

        korean_col = None
        student_start_row = None
        for row in ws.iter_rows(min_row=start_row, max_row=start_row + 10, max_col=TEMPLATE_COLS):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip() == "Korean":
                    korean_col = cell.column
                    student_start_row = cell.row + 1
                    break
            if korean_col:
                break

        if korean_col:
            for i, name in enumerate(students):
                ws.cell(row=student_start_row + i, column=korean_col, value=name)

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        block_idx = 0
        max_row = ws.max_row
        while block_idx * TEMPLATE_ROWS + 1 <= max_row:
            start_row = block_idx * TEMPLATE_ROWS + 1
            end_row = start_row + TEMPLATE_ROWS - 1

            abc_found = False
            for row in ws.iter_rows(min_row=start_row, max_row=end_row, max_col=TEMPLATE_COLS):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip() == "담임 강사: ABC":
                        abc_found = True
                        break
                if abc_found:
                    break

            if abc_found:
                to_unmerge = []
                for m in ws.merged_cells.ranges:
                    if m.min_row >= start_row and m.max_row <= end_row:
                        to_unmerge.append(str(m))
                for m in to_unmerge:
                    ws.unmerge_cells(m)
                ws.delete_rows(start_row, TEMPLATE_ROWS)
                max_row -= TEMPLATE_ROWS
            else:
                block_idx += 1

    if "ABC" in wb.sheetnames:
        del wb["ABC"]

    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream

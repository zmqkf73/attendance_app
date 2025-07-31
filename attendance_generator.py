import re
import calendar
import holidays
from io import BytesIO
from datetime import datetime, date
from openpyxl import load_workbook
from openpyxl.styles import Font

def extract_duration_from_comment(cell):
    if cell.comment is None:
        return None
    lines = cell.comment.text.strip().splitlines()
    for line in reversed(lines):
        if any(token in line for token in ['/', '-', '개월']):
            return line.strip()
    return None

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
    return re.sub(r"([\w가-힣])(?=[\(\[\{])", r"\1 ", text)

def replace_tilde_with_dash(text):
    return text.replace("~", "-")

def insert_space_between_adjacent_brackets(text):
    return re.sub(r"(\))(?=\()", r"\1 ", text)

def capitalize_first_word_if_english(text):
    match = re.match(r"^([A-Za-z]+)(\s|$)", text)
    if match:
        first = match.group(1)
        rest = text[len(first):]
        return first.capitalize() + rest
    return text

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

def generate_attendance(
    records,
    template_path,
    year=None,
    month=None,
    day_type="주중",
    manual_holidays=None,
    manual_includes=None
):
    manual_holidays = set(manual_holidays or [])
    manual_includes = set(manual_includes or [])

    TEMPLATE_ROWS = 31
    TEMPLATE_COLS = 32
    used_block_count = {}

    if isinstance(template_path, str):
        wb = load_workbook(template_path, data_only=False, keep_links=False, keep_vba=False, keep_comments=True)
    elif hasattr(template_path, 'read'):
        wb = load_workbook(template_path, data_only=False, keep_links=False, keep_vba=False, keep_comments=True)
    else:
        raise TypeError("template_path must be a file path or a file-like object")

    template_ws = wb["ABC"]

    today = datetime.today()
    used_year = year or today.year
    used_month = month or today.month

    days_kor = ["월", "화", "수", "목", "금", "토", "일"]
    _, last_day = calendar.monthrange(used_year, used_month)

    kr_holidays = holidays.KR(years=used_year)

    for d in manual_holidays:
        kr_holidays[d] = "사용자 지정 공휴일"

    for d in manual_includes:
        if d in kr_holidays:
            del kr_holidays[d]

    valid_dates = []
    for day in range(1, last_day + 1):
        date_obj = datetime(used_year, used_month, day)
        weekday = date_obj.weekday()
        if ((day_type == "주중" and weekday < 5) or (day_type == "토요일" and weekday == 5)):
            if date_obj.date() not in kr_holidays:
                valid_dates.append((days_kor[weekday], date_obj.day))

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

        year_str = str(used_year)[2:]
        month_str = str(used_month)
        formatted_date = f"{year_str}년 {month_str}월"

        ws.cell(row=start_row + 2, column=2).value = formatted_date
        ws.cell(row=start_row + 2, column=7).value = f"{day} {time}"

        for i in range(23):
            col_idx = 7 + i
            weekday_cell = ws.cell(row=start_row + 4, column=col_idx)
            date_cell = ws.cell(row=start_row + 5, column=col_idx)

            if i < len(valid_dates):
                weekday_cell.value = valid_dates[i][0] 
                date_cell.value = valid_dates[i][1]    
            else:
                weekday_cell.value = None
                date_cell.value = None

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
            duration_col = None
            for row in ws.iter_rows(min_row=start_row, max_row=start_row + 10, max_col=TEMPLATE_COLS):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip() == "수강기간":
                        duration_col = cell.column
                        break
                if duration_col:
                    break
            for i, student_dict in enumerate(students):
                name = student_dict.get("name")
                if not name:
                    continue
                row_idx = student_start_row + i
                cell = ws.cell(row=row_idx, column=korean_col)
                cell.value = name

                duration = extract_duration_from_comment(cell)
                if duration:
                    ws.cell(row=row_idx, column=duration_col).value = duration

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

    if "ABC" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["ABC"]

    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
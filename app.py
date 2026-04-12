import os
import re
import tempfile
from calendar import monthrange
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from attendance_generator import generate_attendance
from attendance_parser import SHEET_CONFIGS, normalize_teacher_name, parse_language_records


def _find_column(columns, keywords, exclude=()):
    for column in columns:
        label = str(column).strip()
        if any(keyword.lower() in label.lower() for keyword in keywords):
            if not any(ex.lower() in label.lower() for ex in exclude):
                return column
    return None


def _clean_teacher_option(value):
    if value is None:
        return None

    teacher = normalize_teacher_name(str(value)).strip()
    if not teacher or teacher.lower() == "nan" or teacher == "강사":
        return None
    if teacher.lower() in {"select all", "all", "전체", "전체선택", "전체 선택"}:
        return None
    if teacher.startswith("※"):
        return None
    if re.search(r"\d", teacher):
        return None
    if any(keyword in teacher for keyword in ("수업", "예정", "시험", "대비반", "신규", "변경")):
        return None

    return teacher


def load_teacher_options(workbook_path):
    teacher_names = set()

    for sheet_name, header_row, _, _ in SHEET_CONFIGS:
        try:
            df = pd.read_excel(
                workbook_path,
                sheet_name=sheet_name,
                header=header_row - 1,
                engine="openpyxl",
            )
        except Exception:
            continue

        df.columns = [
            column if isinstance(column, (int, float)) else str(column).strip()
            for column in df.columns
        ]

        teacher_col = _find_column(df.columns, ["강사"], exclude=["인원", "보"])
        if teacher_col is None:
            continue

        values = (
            df[teacher_col]
            .fillna("")
            .astype(str)
            .map(str.strip)
            .tolist()
        )
        teacher_names.update(
            teacher
            for teacher in (_clean_teacher_option(value) for value in values)
            if teacher
        )

    return sorted(teacher_names)


st.set_page_config(page_title="출석부 생성기", layout="centered")
st.title("출석부 자동 생성기")
st.markdown("업무용 시간표 엑셀 파일을 업로드하고 출석부를 생성하세요.")
st.markdown("출석부 양식 템플릿은 내부에 포함된 `template.xlsx` 파일을 사용합니다.")

uploaded_file = st.file_uploader("시간표 엑셀 파일 업로드", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    selected_year = st.selectbox("년도", options=list(range(2024, 2031)))
with col2:
    selected_month = st.selectbox("월", options=list(range(1, 13)))

selected_day_type = st.radio("출석 요일 유형 선택", options=["주중", "토요일"], index=0)

with st.expander("사용자 지정 날짜 설정"):
    _, last_day = monthrange(selected_year, selected_month)
    day_options = [f"{selected_year}-{selected_month:02d}-{d:02d}" for d in range(1, last_day + 1)]

    manual_holiday_strs = st.multiselect("추가로 제외할 날짜 선택", options=day_options, default=[])
    manual_include_strs = st.multiselect("추가로 포함할 날짜 선택", options=day_options, default=[])

    def parse_dates(date_strs):
        result = []
        for s in date_strs:
            try:
                result.append(datetime.strptime(s, "%Y-%m-%d").date())
            except ValueError:
                continue
        return result

    manual_holidays = parse_dates(manual_holiday_strs)
    manual_includes = parse_dates(manual_include_strs)

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        tmp_input.write(uploaded_file.getvalue())
        tmp_input_path = tmp_input.name

    try:
        with st.spinner("강사 목록을 불러오는 중..."):
            all_teachers = load_teacher_options(tmp_input_path)

        if not all_teachers:
            st.error("강사 목록을 찾지 못했습니다. 업로드한 파일 형식을 확인해주세요.")
        else:
            selected_teachers = st.multiselect(
                "출석부를 생성할 강사를 선택하세요 (선택 없으면 전체 생성)",
                all_teachers,
                placeholder="선택하지 않으면 전체 강사 출석부를 생성합니다.",
            )

            generate = st.button("출석부 생성")

            if generate:
                with st.spinner("출석부 생성 중..."):
                    records = parse_language_records(tmp_input_path)

                    target_set = None if not selected_teachers else set(selected_teachers)
                    filtered_records = [
                        record for record in records
                        if target_set is None or record["강사"] in target_set
                    ]

                    base_dir = os.path.dirname(os.path.abspath(__file__))
                    template_path = os.path.join(base_dir, "template.xlsx")
                    if not Path(template_path).exists():
                        raise FileNotFoundError(f"template.xlsx not found at {template_path}")

                    output_stream = generate_attendance(
                        filtered_records,
                        template_path=template_path,
                        year=selected_year,
                        month=selected_month,
                        day_type=selected_day_type,
                        manual_holidays=manual_holidays,
                        manual_includes=manual_includes,
                    )

                filename = f"{selected_year}년_{selected_month:02d}월_출석부.xlsx"
                st.success("출석부 생성이 완료되었습니다.")
                st.download_button(
                    "출석부 다운로드",
                    data=output_stream.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    finally:
        try:
            os.unlink(tmp_input_path)
        except OSError:
            pass

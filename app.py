import os
import tempfile
from calendar import monthrange
from datetime import datetime
from pathlib import Path

import streamlit as st

from attendance_generator import generate_attendance
from attendance_parser import parse_language_records


st.set_page_config(page_title="출석부 생성기", layout="centered")
st.title("출석부 자동 생성기")
st.markdown("학생 명단 파일을 업로드하면 언어 시트(영어/일본어/중국어/한국어)만 읽어서 출석부를 만듭니다.")
st.markdown("출석부 양식은 같은 폴더의 `template.xlsx`를 사용합니다.")

uploaded_file = st.file_uploader("학생 명단 파일 업로드", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    selected_year = st.selectbox("연도", options=list(range(2024, 2031)))
with col2:
    selected_month = st.selectbox("월", options=list(range(1, 13)))

selected_day_type = st.radio("출석 요일 유형", options=["주중", "토요일"], index=0)

with st.expander("공휴일/포함 날짜 설정"):
    _, last_day = monthrange(selected_year, selected_month)
    day_options = [f"{selected_year}-{selected_month:02d}-{day:02d}" for day in range(1, last_day + 1)]

    manual_holiday_strs = st.multiselect("추가로 제외할 날짜", options=day_options, default=[])
    manual_include_strs = st.multiselect("추가로 포함할 날짜", options=day_options, default=[])

    def parse_dates(date_strs):
        result = []
        for date_str in date_strs:
            try:
                result.append(datetime.strptime(date_str, "%Y-%m-%d").date())
            except ValueError:
                continue
        return result

    manual_holidays = parse_dates(manual_holiday_strs)
    manual_includes = parse_dates(manual_include_strs)

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        tmp_input.write(uploaded_file.read())
        tmp_input_path = tmp_input.name

    try:
        records = parse_language_records(tmp_input_path)
        all_teachers = sorted({record["강사"] for record in records if record.get("강사")})

        selected_teachers = st.multiselect(
            "출석부를 생성할 강사 선택 (선택 없으면 전체 생성)",
            all_teachers,
        )

        generate = st.button("출석부 생성")

        if generate:
            target_set = None if not selected_teachers else set(selected_teachers)
            filtered_records = [
                record for record in records
                if target_set is None or record["강사"] in target_set
            ]

            base_dir = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_dir, "template.xlsx")
            if not Path(template_path).exists():
                raise FileNotFoundError(f"template.xlsx not found at {template_path}")

            with st.spinner("출석부 생성 중..."):
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

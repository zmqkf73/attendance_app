import streamlit as st
import tempfile
import os
import pandas as pd
from datetime import datetime
from calendar import monthrange
from pathlib import Path
from attendance_generator import (
    generate_attendance,
    format_text,
    capitalize_first_word_if_english,
    clean_name,
)

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
    st.markdown("- 아래에서 해당 월의 일(day)을 선택하세요.")

    _, last_day = monthrange(selected_year, selected_month)
    day_options = [f"{selected_year}-{selected_month:02d}-{d:02d}" for d in range(1, last_day + 1)]

    manual_holiday_strs = st.multiselect(
        "추가로 제외할 날짜 선택",
        options=day_options,
        default=[]
    )

    manual_include_strs = st.multiselect(
        "추가로 포함할 날짜 선택",
        options=day_options,
        default=[]
    )

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
        tmp_input.write(uploaded_file.read())
        tmp_input_path = tmp_input.name

    df = pd.read_excel(tmp_input_path, header=5)
    df.columns = [str(c).strip() for c in df.columns]  # ← 반드시 먼저 수행

    category_col = "구분"
    course_col = "과정"
    day_col = "요일"
    time_col = "시간"
    teacher_col = "강사"
    material_col = "교재"
    student_end_col = "총인원"

    df[category_col] = df[category_col].fillna(method="ffill")
    st.write("df.columns:", list(df.columns))

    start_index = df.columns.get_loc(material_col) + 1
    end_index = df.columns.get_loc(student_end_col)
    student_cols = df.columns[start_index:end_index]

    df[teacher_col] = df[teacher_col].fillna("").astype(str).map(str.strip)
    all_teachers = sorted(set(
        capitalize_first_word_if_english(format_text(t))
        for t in df[teacher_col].unique()
        if t and t.lower() != "nan" and t != "강사"
    ))

    selected_teachers = st.multiselect("출석부를 생성할 강사를 선택하세요 (선택 없으면 전체 생성)", all_teachers)

    generate = st.button("출석부 생성")

    if generate:
        y = selected_year
        m = selected_month
        target_set = None if not selected_teachers else set(selected_teachers)

        records = []
        for _, row in df.iterrows():
            teacher_raw = row.get(teacher_col)
            teacher = capitalize_first_word_if_english(format_text(teacher_raw))

            if target_set is None or teacher in target_set:
                category = format_text(row.get(category_col))
                course = format_text(row.get(course_col))
                day = format_text(row.get(day_col))
                time = format_text(row.get(time_col))

                students = []
                for col in student_cols:
                    name_raw = row[col]
                    if pd.isna(name_raw):
                        continue
                    name = clean_name(format_text(str(name_raw).strip()))
                    if name:
                        students.append({"name": name})

                records.append({
                    "구분": category,
                    "과정": course,
                    "요일": day,
                    "시간": time,
                    "강사": teacher,
                    "학생목록": students
                })

        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "template.xlsx")
        if not Path(template_path).exists():
            raise FileNotFoundError(f"template.xlsx not found at {template_path}")
        
        st.write("template_path:", template_path)
        st.write("File exists:", os.path.exists(template_path))

        with st.spinner("출석부 생성 중..."):
            output_stream = generate_attendance(
                records,
                template_path=template_path,
                year=y,
                month=m,
                day_type=selected_day_type,
                manual_holidays=manual_holidays,
                manual_includes=manual_includes
            )

        filename = f"{y}년_{m:02d}월_출석부.xlsx"
        st.success("출석부 생성이 완료되었습니다.")
        st.download_button(
            "출석부 다운로드",
            data=output_stream.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
import streamlit as st
import tempfile
import os
import pandas as pd
from datetime import datetime
from attendance_generator import generate_attendance, format_text, capitalize_first_word_if_english, clean_name

st.set_page_config(page_title="출석부 생성기", layout="centered")

st.title("출석부 자동 생성기")
st.markdown("업무용 시간표 엑셀 파일을 업로드하고 출석부를 생성하세요.")
st.markdown("출석부 양식 템플릿은 내부에 포함된 `template.xlsx` 파일을 사용합니다.")

uploaded_file = st.file_uploader("시간표 엑셀 파일 업로드", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    selected_year = st.selectbox("년도", options=["선택 안 함"] + list(range(2024, 2031)), index=0)
with col2:
    selected_month = st.selectbox("월", options=["선택 안 함"] + list(range(1, 13)), index=0)

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        tmp_input.write(uploaded_file.read())
        tmp_input_path = tmp_input.name

    df = pd.read_excel(tmp_input_path, header=5)

    category_col = "구분"
    course_col = "과정"
    day_col = "요일"
    time_col = "시간"
    teacher_col = "강사"
    material_col = "교재"
    student_end_col = "총인원"

    df.columns = [str(c).strip() for c in df.columns]
    df[category_col] = df[category_col].fillna(method="ffill")

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
        y = int(selected_year) if selected_year != "선택 안 함" else None
        m = int(selected_month) if selected_month != "선택 안 함" else None
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
                students = row[student_cols].dropna().astype(str).apply(clean_name).tolist()

                records.append({
                    "구분": category,
                    "과정": course,
                    "요일": day,
                    "시간": time,
                    "강사": teacher,
                    "학생목록": sorted(set(students))
                })

        template_path = os.path.join(os.path.dirname(__file__), "template.xlsx")

        with st.spinner("출석부 생성 중..."):
            output_stream = generate_attendance(records, template_path, year=y, month=m)

        filename = f"{y or datetime.today().year}년_{m or datetime.today().month:02d}월_출석부.xlsx"
        st.success("출석부 생성이 완료되었습니다.")
        st.download_button(
            "출석부 다운로드",
            data=output_stream.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

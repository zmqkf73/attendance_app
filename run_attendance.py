#!/usr/bin/env python3
"""2025년 8월 출석부 생성 스크립트."""

import sys
from pathlib import Path

sys.path.insert(0, ".")

from attendance_generator import generate_attendance  # noqa: E402
from attendance_parser import SHEET_CONFIGS, parse_sheet  # noqa: E402


STUDENT_LIST = "2025.8월 시간표 학생명단.xlsx"
TEMPLATE = "template.xlsx"
OUTPUT_DIR = Path("output")
YEAR, MONTH = 2025, 8


all_records = []
for sheet_name, header_row, day_col_idx, preferred_course_col in SHEET_CONFIGS:
    print(f"\n[{sheet_name}] 파싱 중...")
    records = parse_sheet(
        STUDENT_LIST,
        sheet_name,
        header_row,
        day_col_idx,
        preferred_course_col,
    )
    print(f"  수업 수: {len(records)}")
    for record in records:
        print(
            f"    강사={record['강사']} | 과정={record['과정'][:20]} "
            f"| 요일={record['요일'][:15]} | 시간={record['시간']} "
            f"| 학생={len(record['학생목록'])}명"
        )
    all_records.extend(records)

print(f"\n총 수업 수: {len(all_records)}")
print("출석부 생성 중...")

output_stream = generate_attendance(
    all_records,
    template_path=TEMPLATE,
    year=YEAR,
    month=MONTH,
    day_type="주중",
)

OUTPUT_DIR.mkdir(exist_ok=True)
out_path = OUTPUT_DIR / f"{YEAR}년_{MONTH:02d}월_출석부.xlsx"
with open(out_path, "wb") as output_file:
    output_file.write(output_stream.getvalue())

print(f"저장 완료: {out_path}")

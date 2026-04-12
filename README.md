# Attendance App

**Live App**: [https://attendanceapp-kgm2r8ahpkesw4wpmrxxsq.streamlit.app/](https://attendanceapp-kgm2r8ahpkesw4wpmrxxsq.streamlit.app/)

This project is a Streamlit-based attendance sheet generator for internal use.

## What It Does

- Upload a student schedule workbook in `.xlsx` format
- Read only the language sheets: `영어`, `일본어`, `중국어`, `한국어`
- Group classes by instructor
- Generate attendance sheets from `template.xlsx`
- Download a single compiled workbook with a filename like `2025년_08월_출석부.xlsx`

## Current Behavior

- The app keeps the template layout as-is and fills in attendance data
- Long `CLASS:` values are inserted without changing the template width
- Instructor aliases such as `Ray`, `윤정원 (레이)`, and `Ray (윤정원)` are normalized to `Ray (윤정원)`
- Non-student labels such as `일대일`, `일대일 수업`, and similar class text are filtered out from the student list

## How to Run Locally

1. Clone the repository:

```bash
git clone https://github.com/zmqkf73/attendance_app.git
cd attendance_app
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Launch the app:

```bash
streamlit run app.py
```

## Input

- Upload a `.xlsx` workbook containing the student schedule / roster
- The workbook should include the language sheets used by the app
- The parser currently reads:
  - `영어`: header row 6, course column `과정`
  - `일본어`: header row 7, day column 5, course column `과정.1`
  - `중국어`: header row 6, course column `과정1`
  - `한국어`: header row 6, course column `과정1`

## Template

- Attendance sheets are generated from `template.xlsx`
- `template.xlsx` must stay in the project root next to `app.py`
- The template file should be committed to the repository

## Output

- The app returns one compiled workbook containing attendance sheets by instructor
- The download filename format is:

```text
YYYY년_MM월_출석부.xlsx
```

Example:

```text
2025년_08월_출석부.xlsx
```

## Main Files

- `app.py`: Streamlit UI
- `attendance_parser.py`: schedule parsing and normalization
- `attendance_generator.py`: workbook generation from the template
- `run_attendance.py`: local script runner for direct generation

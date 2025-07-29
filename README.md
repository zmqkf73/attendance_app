# Attendance App

This is a Streamlit-based attendance sheet generator for internal use.

## Features
- Upload an internal Excel schedule file
- Select instructor(s), year, and month
- Choose weekday or Saturday attendance format
- Automatically generate formatted attendance sheets using a pre-defined template

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

## Input Format
- Upload a `.xlsx` file containing the internal schedule
- The file should match the structure expected by the app (starting from row 6)

## Template
- The attendance sheets are generated based on `template.xlsx`
- This template must exist in the same directory as `app.py`

## Output
- Attendance files are generated per instructor
- Users can download the final compiled `.xlsx` file via the app interface

## Notes
- Temporary Excel files (e.g., `~$template.xlsx`) are ignored via `.gitignore`
- Only weekdays (Mon–Fri) or Saturdays are populated depending on selection

---
For internal use only. Make sure to keep any uploaded schedule files secure.

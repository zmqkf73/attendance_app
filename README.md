# Attendance App

**Live App**: [https://attendanceapp-kgm2r8ahpkesw4wpmrxxsq.streamlit.app/](https://attendanceapp-kgm2r8ahpkesw4wpmrxxsq.streamlit.app/)

Attendance App is a Streamlit-based tool for generating attendance workbooks from an uploaded Excel file and a predefined template.

## Features

- Upload an Excel workbook
- Select year and month
- Generate formatted attendance sheets automatically
- Download the final workbook as an `.xlsx` file

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

- Upload a `.xlsx` workbook in the format expected by the application

## Template

- Attendance sheets are generated from `template.xlsx`
- `template.xlsx` must be located in the project root

## Output

- The app generates a downloadable Excel workbook
- Output filenames follow this format:

```text
YYYY년_MM월_출석부.xlsx
```

Example:

```text
2025년_08월_출석부.xlsx
```

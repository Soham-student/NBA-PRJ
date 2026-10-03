# Course Outcome Mapping and Attainment Workbook

This project extracts and organizes CO, PO, PSO, and attainment data into an Excel workbook.

## Overview

The workflow is:

1. Place the input Excel file in the `input_sheets` folder.
2. Run the Python script.
3. The script reads the workbook and extracts the relevant tables.
4. It creates a formatted Excel output in the `output_sheets` folder.
5. The output workbook contains:
   - Raw extracted data
   - CO–PO–PSO mapping calculations
   - Course attainment summary

## Folder Structure

```text
project_bundle/
├── extracton.py
├── input_sheets/
│   └── input_sheet.xlsx
├── output_sheets/
│   └── course_attainment_analysis.xlsx
├── README.md
```

## Requirements

Install the required Python package:

```powershell
pip install openpyxl
```

If you are using the project virtual environment, run:

```powershell
.\.venv\Scripts\python.exe -m pip install openpyxl
```

## How to Run

From the project folder:

```powershell
Set-Location "E:\WEB DEV\backend\project_bundle"
.\.venv\Scripts\python.exe .\extracton.py
```

Or, if the virtual environment is already active:

```powershell
python .\extracton.py
```

## Input File

Place your Excel source file in:

```text
project_bundle\input_sheets\
```

The script will look for:

```text
project_bundle\input_sheets\input_sheet.xlsx
```

If that file is not present, it will generate the sample workbook automatically.

## Output File

The generated output is saved to:

```text
project_bundle\output_sheets\course_attainment_analysis.xlsx
```

## Notes

- The script is designed for academic course outcome mapping and attainment reporting.
- It supports importing the CO/PO/PSO tables and creating structured summary sheets.
- The workbook output is formatted for clarity and presentation.

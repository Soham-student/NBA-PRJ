from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


RAW_DATA = {
    "Table 1: Mapping of Cos with Pos and PSOs": [
        ["CO", "PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2", "PSO3"],
        ["CO1", "√", "√", "√", "0", "0", "0", "0", "√", "", "√", "0", "√", "", "√", ""],
        ["CO2", "√", "√", "√", "0", "0", "0", "0", "√", "", "√", "0", "√", "", "√", ""],
        ["CO3", "√", "√", "√", "0", "√", "0", "0", "√", "", "√", "0", "√", "", "√", ""],
        ["CO4", "√", "√", "√", "0", "0", "0", "0", "√", "", "√", "0", "√", "", "√", ""],
        ["CO5", "√", "√", "√", "0", "0", "0", "0", "√", "", "√", "0", "√", "", "√", ""],
    ],
    "Table 2: Instructional hours alloted to each unit and their mapping with COs": [
        ["Column1", "Instructional hrs (Theory)", "CO Mapped", "PO/PSO- Mapped", "Column2", "Column3", "Column4", "Column5", "Assignment No.", "Name of the assignment", "Unit", "CO Mapped", "CO", "Intructional Hrs Theory", "Intructional Hrs Practical", "Total Instructional Hrs", "Mapping Strenght", "Reference Level"],
        ["Unit-1", "9", "CO1", "PO1, PO2, PO8, PO9, PO10,PO12,PSO1,PSO2,PSO3", "", "", "", "", "1", "Assembler Pass 1", "1", "CO1", "CO1", "9", "6", "15", "H=3", "H", ">30"],
        ["Unit-2", "9", "CO2", "PO1, PO2, PO8, PO9, PO10 ,PO12,PSO1,PSO2,PSO3", "", "", "", "", "2", "Assembler Pass 2", "1", "CO1", "CO2", "9", "8", "17", "H=3", "M", "16 to 29"],
        ["Unit-3", "9", "CO3", "PO1,PO2, PO3, PO8, PO9, PO10 ,PO12,PSO1,PSO2,PSO3", "", "", "", "", "3", "Macro Pass 1", "2", "CO2", "CO3", "9", "16", "25", "H=3", "L", "5 to15"],
        ["Unit-4", "9", "CO4", "PO2, PO5, PO8, PO9, PO10,PO12,PSO1,PSO2,PSO3", "", "", "", "", "4", "Macro Pass 2", "2", "CO2", "CO4", "9", "0", "9", "M=2", "", "0", "<5"],
        ["Unit-5", "9", "CO5", "PO1,PO2,PO3,PO12,PSO1,PSO2,PSO3", "", "", "", "", "5", "Lexical Analysis", "3", "CO3", "CO5", "9", "0", "9", "M=2", "", ""],
        ["Total", "45", "", "", "", "", "", "", "6", "RDP", "3", "CO3", "", "45", "30", "75", "", ""],
        ["", "", "", "", "", "", "", "", "7", "Compund statement", "3", "CO3", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "8", "calculator using lex and yacc", "3", "CO3", "", "", "", "", "", ""],
    ],
    "Table 7: Final CO to PO PSO mapping": [
        ["CO", "PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2", "PSO3"],
        ["CO1", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO2", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO3", 3, 3, 3, 0, 3, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO4", 3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        ["CO5", 3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        ["Target attainment", 3, 3, 3, 0, 3, 0, 0, 2.6, 2.6, 2.6, 0, 2.6, 2.6, 2.6, 2.6],
    ],
    "Units and assessment summary": [
        ["Units", "CO Mapped", "PO/PSO- Mapped", "CCA", "LCA", "End-Term"],
        ["Unit-1", "CO1", "PO1, PO2, PO8, PO9, PO10,PO12,PSO1,PSO2,PSO3", 30, 30, 40],
        ["Unit-2", "CO2", "PO1, PO2, PO8, PO9, PO10 ,PO12,PSO1,PSO2,PSO3", 60, "", ""],
        ["Unit-3", "CO3", "PO1,PO2, PO3,PO5, PO8, PO9, PO10 ,PO12,PSO1,PSO2,PSO3", "", "", ""],
        ["Unit-4", "CO4", "PO1,PO2, PO3,, PO8, PO9, PO10,PO12,PSO1,PSO2,PSO3", "", "", ""],
        ["Unit 5", "CO5", "PO1,PO2, PO3,, PO8, PO9, PO10,PO12,PSO1,PSO2,PSO3", "", "", ""],
    ],
    "CCA Distribution": [
        ["", "Assignment 1 (/10)", "Theory assignment(/5)", "MID TERM (/15)", "Total"],
        ["Actual marks", 10, 5, 15, 30],
        ["% Marks", 33.33, 16.67, 50, 100.00],
    ],
    "Mapping with Lab assignments": [
        ["Assignment No.", "Distribution of Practical assignments", "CO Mapped", "Weight for each assignment"],
        [1, "Assembler Pass 1", "CO1", "1/8"],
        [2, "Assembler Pass 2", "CO1", "1/8"],
        [3, "Macro Pass 1", "CO2", "1/8"],
        [4, "Macro Pass 2", "CO2", "1/8"],
        [5, "Lexical Analysis", "CO3", "1/8"],
        [6, "RDP", "CO3", "1/8"],
        [7, "Compund statement", "CO3", "1/8"],
        [8, "calculator using lex and yacc", "CO3", "1/8"],
    ],
    "CO contribution summary": [
        ["CO", "Fraction of CO mapped", "Marks out of 50"],
        ["CO1", "1/8+1/8", "1/4", 7.5],
        ["CO2", "1/8+1/8", "1/4", 7.5],
        ["CO3", "1/8+1/8+1/8+1/8", "1/2", 15.0],
        ["CO4", "", "", ""],
        ["CO5", "", "", ""],
    ],
    "Actual value wise distribution of tools": [
        ["", "Assignment 1 (/10)", "Active Learning", "MID TERM (/15)", "", "LCA", "SUM"],
        ["CO1", 3, 0, 7, "", 7.5, 17.5],
        ["CO2", 2, "", 8, "", 7.5, 17.5],
        ["CO3", 0, 5, "", "", 15.0, 20],
        ["CO4", "", 2, "", "", "", 2],
        ["CO5", "", 3, "", "", "", 3],
    ],
    "%Contribution of individual tool for each CO": [
        ["", "Assignment 1 (/10)", "PBL", "MID TERM (/15)", "MCQ (/10)", "Attendance (/5)", "LCA"],
        ["CO1", 17.14, 0, 40, 0, 0, 42.86],
        ["CO2", 11.43, 0, 45.71, 0, 0, 42.86],
        ["CO3", 0, 25, 0, 0, 0, 75],
        ["CO4", 0, 100, 0, 0, 0, 0],
        ["CO5", 0, 100, 0, 0, 0, 0],
    ],
    "Course attainment summary": [
        ["Table A", "", "", "", ""],
        ["Rules for finding CO level (* Should be decided by Course Coordinators along with the team members)", "", "", "", ""],
        ["Course Outcome", "Target attainment (As per form 8 of course file)", "3", "2", "1"],
        ["CO1", 60, ">=60%", "50-59%", "40-49%"],
        ["CO2", 50, ">=60%", "50-59%", "40-49%"],
        ["CO3", 55, ">=60%", "50-59%", "40-49%"],
        ["CO4", 50, ">=60%", "50-59%", "40-49%"],
        ["CO5", 50, ">=60%", "50-59%", "40-49%"],
        ["", "", "", "", ""],
        ["Table B", "", "", "", ""],
        ["CO Attainment (Direct) Gap analysis", "", "", "", ""],
        ["Course Outcome", "% Direct attainment", "Actual % Indirect attainment (% of students saying 'YES' in End course survey)", "Total actual CO attainment", "CO attainment level", "IA CO Leve;", "Status of attainment (Attained /Not attained)"],
        ["CO1", 60.85346939, 80, 64.68277551, 3, 3, "Attained"],
        ["CO2", 48.54, 80, 54.832, 3, 3, "Attained"],
        ["CO3", 52.43, 80, 57.944, 3, 3, "Attained"],
        ["CO4", 40.47, 85, 49.376, 2, 3, "not attained"],
        ["CO5", 47.45, 80, 53.96, 3, 3, "Attained"],
        ["Average CO Attainment", 2.84, "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["Table F", "", "", "", "", "", "", ""],
        ["Overall attainment for course over panel B", "", "", "", "", "", "", ""],
        ["CO", "PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2", "PSO3"],
        ["CO1", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO2", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO3", 3, 3, 3, 0, 3, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO4", 3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        ["CO5", 3, 3, 3, 2, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        ["Target attainment", 3, 3, 3, 2, 3, 0, 0, 2.6, 2.6, 2.6, 0, 2.6, 2.71, 2.63, 2.6],
        ["Actual attainment", 2.84, 2.84, 2.84, 1.89, 2.84, 0, 0, 2.46, 2.46, 2.46, 0, 2.46, 2.57, 2.49, 2.46],
        ["Note: For finding average target (expected ) of each PO consider only the cells having non zero entries.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["E.g. for PO1,", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["Target Attainment = (2+2+2)/3 = 2", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["To find actual attainment,multiply the entry in target attainment by value obtained for Overall attainment of COs(in Table E) and divide by maximum value that is 3.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["E.g. for PO1,", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["Actual attainment = 2*2.53/3 =1.69", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
    ],
    "%Contribution of individual tool for each CO (Sheet 3 summary)": [
        ["", "Assignment 1 (/5)", "PBL(/10)", "MID TERM (/15)", "", "LCA"],
        ["CO1", 17.14, 0, 40, 0, 42.86, 100],
        ["CO2", 11.43, 0, 45.71, 0, 42.86, 100.01],
        ["CO3", 0, 25, 0, 0, 75, 100],
        ["CO4", 0, 100, 0, 0, 0, 100],
        ["CO5", 0, 100, 0, 0, 0, 100],
    ],
}

HEADERS = [
    "CO",
    "PO1",
    "PO2",
    "PO3",
    "PO4",
    "PO5",
    "PO6",
    "PO7",
    "PO8",
    "PO9",
    "PO10",
    "PO11",
    "PO12",
    "PSO1",
    "PSO2",
    "PSO3",
]


def normalize_cell(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def detect_columns(headers: Sequence[str]) -> dict:
    mapping = {}
    for idx, header in enumerate(headers):
        key = normalize_header(header)
        if key:
            mapping[key] = idx
    return mapping


def write_rows(worksheet, rows: Iterable[Sequence], start_row: int = 1):
    for row_index, row in enumerate(rows, start=start_row):
        for col_index, value in enumerate(row, start=1):
            worksheet.cell(row=row_index, column=col_index, value=value)


def extract_tables_from_excel(file_path: str | Path) -> dict:
    workbook = load_workbook(file_path, data_only=True)
    extracted = {}
    for sheet in workbook.sheetnames:
        ws = workbook[sheet]
        rows = list(ws.iter_rows(values_only=True))
        for index, row in enumerate(rows):
            clean_row = [normalize_cell(v) for v in row]
            normalized = [normalize_header(v) for v in clean_row if v]
            if "co" in normalized or any(item.startswith("po") or item.startswith("pso") for item in normalized):
                header = clean_row
                table_rows = []
                for next_row in rows[index + 1 :]:
                    if all(normalize_cell(v) == "" for v in next_row):
                        break
                    table_rows.append([normalize_cell(v) for v in next_row])
                extracted[f"{sheet}:{index + 1}"] = [header] + table_rows
                break
    return extracted


def style_sheet(ws):
    thin = Border(left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"), top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    title_fill = PatternFill("solid", fgColor="CFE2F3")
    header_font = Font(bold=True)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center")

    if ws.max_row > 0:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str):
                text = str(cell.value).strip()
                if text.lower().startswith("table ") or text.lower().startswith("overall"):
                    cell.fill = title_fill
                    cell.font = Font(bold=True, color="000000")

    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]:
        if col_letter in ws.column_dimensions:
            ws.column_dimensions[col_letter].width = 15
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18

    # Format rows with Target Attainment and Actual Attainment to match the sample style
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "Target attainment" in cell.value:
                for c in ws[row[0].row]:
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
                    c.font = Font(bold=True)
            if isinstance(cell.value, str) and "Actual attainment" in cell.value:
                for c in ws[row[0].row]:
                    c.fill = PatternFill("solid", fgColor="C6EFCE")
                    c.font = Font(bold=True)


def build_raw_extract_sheet(wb: Workbook, raw_tables: dict | None = None):
    ws = wb.active
    ws.title = "RawData"
    source_tables = raw_tables if raw_tables is not None else RAW_DATA
    row_cursor = 1
    for table_name, rows in source_tables.items():
        ws.cell(row=row_cursor, column=1, value=table_name)
        ws.cell(row=row_cursor, column=1).font = Font(bold=True, color="000000")
        ws.cell(row=row_cursor, column=1).fill = PatternFill("solid", fgColor="CFE2F3")
        row_cursor += 1

        if not rows:
            row_cursor += 1
            continue

        write_rows(ws, rows, start_row=row_cursor)
        row_cursor += len(rows) + 2

    style_sheet(ws)
    ws.freeze_panes = "A2"


def build_calculation_sheet(wb: Workbook):
    ws = wb.create_sheet(title="Calculations")
    final_mapping = [
        ["CO", "PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2", "PSO3"],
        ["CO1", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO2", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO3", 3, 3, 3, 0, 3, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO4", 3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        ["CO5", 3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        ["Target attainment", 3, 3, 3, 0, 3, 0, 0, 2.6, 2.6, 2.6, 0, 2.6, 2.6, 2.6, 2.6],
    ]

    write_rows(ws, final_mapping, start_row=1)
    ws["A1"].value = "Final CO to PO PSO mapping (From Table 1)"
    ws["A1"].font = Font(bold=True)

    summary_start = 10
    summary_rows = [
        ["Assessment Type", "Marks", "% Contribution"],
        ["CCA", 30, 40],
        ["LCA", 30, 40],
        ["End-Term", 40, 20],
    ]
    write_rows(ws, summary_rows, start_row=summary_start)
    ws[f"A{summary_start}"].font = Font(bold=True)

    unit_rows = [
        ["Units", "CO Mapped", "PO/PSO- Mapped", "CCA", "LCA", "End-Term"],
        ["Unit-1", "CO1", "PO1, PO2, PO8, PO9, PO10, PO12, PSO1, PSO2, PSO3", 30, 30, 40],
        ["Unit-2", "CO2", "PO1, PO2, PO8, PO9, PO10, PO12, PSO1, PSO2, PSO3", 60, "", ""],
        ["Unit-3", "CO3", "PO1, PO2, PO3, PO5, PO8, PO9, PO10, PO12, PSO1, PSO2, PSO3", "", "", ""],
        ["Unit-4", "CO4", "PO1, PO2, PO3, PO8, PO9, PO10, PO12, PSO1, PSO2, PSO3", "", "", ""],
        ["Unit 5", "CO5", "PO1, PO2, PO3, PO8, PO9, PO10, PO12, PSO1, PSO2, PSO3", "", "", ""],
    ]
    write_rows(ws, unit_rows, start_row=summary_start + 6)

    lab_rows = [
        ["CO", "Fraction of CO mapped", "Marks out of 50"],
        ["CO1", "1/8 + 1/8", 7.5],
        ["CO2", "1/8 + 1/8", 7.5],
        ["CO3", "1/8 + 1/8 + 1/8 + 1/8", 15.0],
        ["CO4", "", ""],
        ["CO5", "", ""],
    ]
    write_rows(ws, lab_rows, start_row=summary_start + 15)

    ws["A1"].value = "Final CO to PO PSO mapping (From Table 1)"
    ws["A1"].font = Font(bold=True)
    style_sheet(ws)
    ws.freeze_panes = "A2"


def build_course_attainment_sheet(wb: Workbook):
    ws = wb.create_sheet(title="Course_Attainment")

    table_a = [
        ["Table A", "", "", "", ""],
        ["Rules for finding CO level (* Should be decided by Course Coordinators along with the team members)", "", "", "", ""],
        ["Course Outcome", "Target attainment (As per form 8 of course file)", "3", "2", "1"],
        ["CO1", 60, ">=60%", "50-59%", "40-49%"],
        ["CO2", 50, ">=60%", "50-59%", "40-49%"],
        ["CO3", 55, ">=60%", "50-59%", "40-49%"],
        ["CO4", 50, ">=60%", "50-59%", "40-49%"],
        ["CO5", 50, ">=60%", "50-59%", "40-49%"],
    ]
    write_rows(ws, table_a, start_row=1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    ws["A1"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws["A1"].font = Font(bold=True)
    ws["A2"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws["A2"].font = Font(bold=True)

    table_b = [
        ["", "", "", "", ""],
        ["Table B", "", "", "", ""],
        ["CO Attainment (Direct) Gap analysis", "", "", "", ""],
        ["Course Outcome", "% Direct attainment", "Actual % Indirect attainment (% of students saying 'YES' in End course survey)", "Total actual CO attainment", "CO attainment level", "IA CO Leve;", "Status of attainment (Attained /Not attained)"],
        ["CO1", 60.85346939, 80, 64.68277551, 3, 3, "Attained"],
        ["CO2", 48.54, 80, 54.832, 3, 3, "Attained"],
        ["CO3", 52.43, 80, 57.944, 3, 3, "Attained"],
        ["CO4", 40.47, 85, 49.376, 2, 3, "not attained"],
        ["CO5", 47.45, 80, 53.96, 3, 3, "Attained"],
        ["Average CO Attainment", 2.84, "", "", "", "", ""],
    ]
    write_rows(ws, table_b, start_row=15)
    ws.merge_cells(start_row=16, start_column=1, end_row=16, end_column=5)
    ws["A16"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws["A16"].font = Font(bold=True)
    ws["A17"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws["A17"].font = Font(bold=True)

    table_f = [
        ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["Table F", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["Overall attainment for course over panel B", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["CO", "PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2", "PSO3"],
        ["CO1", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO2", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO3", 3, 3, 3, 0, 3, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        ["CO4", 3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        ["CO5", 3, 3, 3, 2, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        ["Target attainment", 3, 3, 3, 2, 3, 0, 0, 2.6, 2.6, 2.6, 0, 2.6, 2.71, 2.63, 2.6],
        ["Actual attainment", 2.84, 2.84, 2.84, 1.89, 2.84, 0, 0, 2.46, 2.46, 2.46, 0, 2.46, 2.57, 2.49, 2.46],
    ]
    write_rows(ws, table_f, start_row=35)
    ws.merge_cells(start_row=36, start_column=1, end_row=36, end_column=16)
    ws.merge_cells(start_row=37, start_column=1, end_row=37, end_column=16)
    ws["A36"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws["A36"].font = Font(bold=True)
    ws["A37"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws["A37"].font = Font(bold=True)

    tool_rows = [
        ["", "Assignment 1 (/5)", "PBL(/10)", "MID TERM (/15)", "", "LCA"],
        ["CO1", 17.14, 0, 40, 0, 42.86, 100],
        ["CO2", 11.43, 0, 45.71, 0, 42.86, 100.01],
        ["CO3", 0, 25, 0, 0, 75, 100],
        ["CO4", 0, 100, 0, 0, 0, 100],
        ["CO5", 0, 100, 0, 0, 0, 100],
    ]
    write_rows(ws, tool_rows, start_row=60)

    note_rows = [
        ["Note: For finding average target (expected ) of each PO consider only the cells having non zero entries."],
        ["E.g. for PO1,"],
        ["Target Attainment = (2+2+2)/3 = 2"],
        ["To find actual attainment,multiply the entry in target attainment by value obtained for Overall attainment of COs(in Table E) and divide by maximum value that is 3."],
        ["E.g. for PO1,"],
        ["Actual attainment = 2*2.53/3 =1.69"],
    ]
    write_rows(ws, note_rows, start_row=70)

    style_sheet(ws)
    ws.freeze_panes = "A2"


def create_excel_report(output_path: str | Path, source_excel: str | Path | None = None):
    output_file = Path(output_path)
    wb = Workbook()
    if source_excel is not None:
        extracted = extract_tables_from_excel(source_excel)
        if extracted:
            build_raw_extract_sheet(wb, raw_tables=extracted)
        else:
            build_raw_extract_sheet(wb, raw_tables=RAW_DATA)
    else:
        build_raw_extract_sheet(wb, raw_tables=RAW_DATA)
    build_calculation_sheet(wb)
    build_course_attainment_sheet(wb)

    save_targets = [output_file]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_targets.append(output_file.with_name(f"{output_file.stem}_{timestamp}.xlsx"))

    last_error = None
    for target in save_targets:
        try:
            wb.save(target)
            return target
        except PermissionError as exc:
            last_error = exc
            continue

    raise last_error


if __name__ == "__main__":
    script_dir = Path(__file__).resolve().parent
    if script_dir.name == "pipeline":
        project_root = script_dir.parent / "project_bundle"
    else:
        project_root = script_dir

    input_dir = project_root / "input_sheets"
    output_dir = project_root / "output_sheets"
    output_dir.mkdir(exist_ok=True)
    input_dir.mkdir(exist_ok=True)

    default_input = input_dir / "input_sheet.xlsx"
    report_path = output_dir / "course_attainment_analysis.xlsx"

    if default_input.exists():
        created = create_excel_report(report_path, source_excel=default_input)
    else:
        created = create_excel_report(report_path)

    print(f"Input folder: {input_dir}")
    print(f"Output folder: {output_dir}")
    print(f"Excel report created at: {created}")

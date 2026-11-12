"""Web frontend for CO-PO attainment: enter course values, upload student marks,
get the computed tables on screen + an Excel report, every run saved to SQLite.

Run:  python app.py   then open http://127.0.0.1:5000
"""
from __future__ import annotations

import json
import math
import re
import sqlite3
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal
from io import BytesIO
from pathlib import Path

from flask import Flask, abort, g, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook, load_workbook

from extracton import style_sheet

ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "nba.db"
DEFAULT_STUDENTS = ROOT / "input_sheets" / "input_sheet.xlsx"

COS = ["CO1", "CO2", "CO3", "CO4", "CO5"]
POS = ["PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2", "PSO3"]
TOOLS = ["Assignment", "PBL", "Mid term", "MCQ", "Att", "LCA"]

# Defaults = the values in input_sheets/input_sheet.xlsx (Tab 1 Table 7, Tab 3 Step-2 + Table A/B).
DEFAULTS = {
    "course_name": "System Software and Compiler Design",
    "course_code": "CET3011B",
    "class_name": "Final Year BTech, CSE (AY 2025-26)",
    "semester": "VIII",
    "mapping": {
        "CO1": [3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        "CO2": [3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        "CO3": [3, 3, 3, 0, 3, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
        "CO4": [3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
        "CO5": [3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
    },
    "max_marks": {co: [5, 10, 15, 0, 0, 30] for co in COS},
    "weights": {  # % contribution of each tool to the CO (Tab 2 / Tab 3 top table)
        "CO1": [17.14, 0, 40, 0, 0, 42.86],
        "CO2": [11.43, 0, 45.71, 0, 0, 42.86],
        "CO3": [0, 25, 0, 0, 0, 75],
        "CO4": [0, 100, 0, 0, 0, 0],
        "CO5": [0, 100, 0, 0, 0, 0],
    },
    "target": {"CO1": 60, "CO2": 50, "CO3": 55, "CO4": 50, "CO5": 50},
    "indirect": {"CO1": 80, "CO2": 80, "CO3": 80, "CO4": 85, "CO5": 80},
    "ia_level": {"CO1": 3, "CO2": 3, "CO3": 3, "CO4": 3, "CO5": 3},
    "level3": 60, "level2": 50, "level1": 40,
    "internal_pct": 60,   # end term gets the rest
    "direct_pct": 80,     # indirect gets the rest
    "end_max": 40,
}


def r2(x: float) -> float:
    """Excel ROUND(x, 2): half away from zero (Python's round() is banker's)."""
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def level_for(pct: float, cfg: dict) -> int:
    if pct >= cfg["level3"]:
        return 3
    if pct >= cfg["level2"]:
        return 2
    if pct >= cfg["level1"]:
        return 1
    return 0


# ---------------------------------------------------------------- calculation

def compute(cfg: dict, students: list[dict]) -> dict:
    """Same chain as 'Tab 3 Attainment calculation', applied uniformly to every student.

    students: [{"sr", "name", "marks": {co: [6 tool marks]}, "end": float}]
    """
    int_f = cfg["internal_pct"] / 100
    end_f = 1 - int_f
    rows = []
    for s in students:
        end_conv = 100 * s["end"] / cfg["end_max"]
        internal, total = {}, {}
        for co in COS:
            # Step-2: mark scaled to the tool's % contribution
            internal[co] = r2(sum(
                r2(m * w / mx) for m, w, mx in zip(s["marks"][co], cfg["weights"][co], cfg["max_marks"][co]) if mx
            ))
            # Step-3: internal% + end-term% split
            total[co] = r2(r2(internal[co] * int_f) + r2(end_conv * end_f))
        rows.append({"sr": s["sr"], "name": s["name"], "end": s["end"], "end_conv": r2(end_conv),
                     "internal": internal, "total": total})

    d_f = cfg["direct_pct"] / 100
    co_rows = []
    for co in COS:
        direct = r2(sum(r["total"][co] for r in rows) / len(rows))
        indirect = cfg["indirect"][co]
        total_att = r2(d_f * direct + (1 - d_f) * indirect)
        level = level_for(total_att, cfg)
        final = r2(d_f * level + (1 - d_f) * cfg["ia_level"][co])
        co_rows.append({"co": co, "target": cfg["target"][co], "direct": direct, "indirect": indirect,
                        "total": total_att, "level": level, "ia_level": cfg["ia_level"][co], "final": final,
                        "status": "Attained" if total_att >= cfg["target"][co] else "Not attained"})
    avg_final = r2(sum(c["final"] for c in co_rows) / len(co_rows))

    po_rows = []
    for i, po in enumerate(POS):
        nonzero = [cfg["mapping"][co][i] for co in COS if cfg["mapping"][co][i]]
        target = r2(sum(nonzero) / len(nonzero)) if nonzero else 0.0
        po_rows.append({"po": po, "target": target, "actual": r2(target * avg_final / 3)})

    return {"students": rows, "cos": co_rows, "avg_final": avg_final, "pos": po_rows}


# ---------------------------------------------------------------- input parsing

def num(raw, label: str, errors: list, lo: float | None = None, hi: float | None = None) -> float:
    try:
        v = float(raw)
        if not math.isfinite(v):
            raise ValueError
    except (TypeError, ValueError):
        errors.append(f"{label}: '{raw}' is not a number")
        return 0.0
    if (lo is not None and v < lo) or (hi is not None and v > hi):
        errors.append(f"{label}: {v:g} must be between {lo:g} and {hi:g}")
    return v


def parse_form(form) -> tuple[dict, list]:
    e: list[str] = []
    cfg = {k: form.get(k, "").strip() for k in ("course_name", "course_code", "class_name", "semester")}
    if not cfg["course_name"]:
        e.append("Course name is required")
    cfg["mapping"] = {co: [num(form.get(f"map_{co}_{po}"), f"Mapping {co}-{po}", e, 0, 3) for po in POS] for co in COS}
    for co in COS:
        for po, x in zip(POS, cfg["mapping"][co]):
            if x != int(x):
                e.append(f"Mapping {co}-{po}: {x:g} must be a whole number 0-3")
        cfg["mapping"][co] = [int(x) for x in cfg["mapping"][co]]
    cfg["max_marks"] = {co: [num(form.get(f"max_{co}_{t}"), f"{co} {t} max marks", e, 0, 1000) for t in TOOLS] for co in COS}
    cfg["weights"] = {co: [num(form.get(f"w_{co}_{t}"), f"{co} {t} weight", e, 0, 100) for t in TOOLS] for co in COS}
    for key, lo, hi in (("target", 0, 100), ("indirect", 0, 100), ("ia_level", 0, 3)):
        cfg[key] = {co: num(form.get(f"{key}_{co}"), f"{co} {key.replace('_', ' ')}", e, lo, hi) for co in COS}
    for key in ("level3", "level2", "level1", "internal_pct", "direct_pct"):
        cfg[key] = num(form.get(key), key, e, 0, 100)
    cfg["end_max"] = num(form.get("end_max"), "End term max marks", e, 1, 1000)

    if not cfg["level3"] > cfg["level2"] > cfg["level1"]:
        e.append("Level thresholds must satisfy level 3 > level 2 > level 1")
    for co in COS:
        for t, w, mx in zip(TOOLS, cfg["weights"][co], cfg["max_marks"][co]):
            if w and not mx:
                e.append(f"{co} {t}: has weight {w:g} but max marks 0")
        total_w = sum(cfg["weights"][co])
        if abs(total_w - 100) > 0.1:
            e.append(f"{co}: tool weights add up to {total_w:g}, must be 100")
    return cfg, e


def _cell_num(v, where: str, errors: list) -> float:
    if v is None or (isinstance(v, str) and not v.strip()):
        return 0.0  # blank = 0, same as the source sheet
    try:
        f = float(v)
        if not math.isfinite(f):
            raise ValueError
        return f
    except (TypeError, ValueError):
        errors.append(f"{where}: '{v}' is not a number")
        return 0.0


def parse_students(data: bytes, cfg: dict) -> tuple[list[dict], list]:
    """Accepts either the original course workbook (sheet 'Tab 3 ...') or the simple template."""
    errors: list[str] = []
    try:
        wb = load_workbook(BytesIO(data), data_only=True)
    except Exception:
        return [], ["Could not read the uploaded file as an .xlsx workbook"]

    tab3 = next((n for n in wb.sheetnames if n.strip().lower().startswith("tab 3")), None)
    students = []
    if tab3:
        ws = wb[tab3]
        rows = list(ws.iter_rows(values_only=True))
        name_hdr = [i for i, r in enumerate(rows) if len(r) > 1 and str(r[1] or "").strip() == "Name of the student"]
        end_hdr = [i for i in name_hdr if len(rows[i]) > 7 and str(rows[i][7] or "").strip().lower().startswith("end term")]
        if not name_hdr or not end_hdr:
            return [], [f"Sheet '{tab3}': could not find the Step-1 marks table and Step-3 End Term column"]
        ends = []
        for r in rows[end_hdr[0] + 1:]:
            if not str(r[1] or "").strip():
                break
            ends.append((str(r[1]).strip(), r[7]))
        for k, r in enumerate(rows[name_hdr[0] + 1:]):
            name = str(r[1] or "").strip()
            if not name:
                break
            if k >= len(ends) or ends[k][0] != name:
                errors.append(f"Student '{name}': not found in the same position in the End Term table")
                continue
            vals = list(r[2:32]) + [None] * (30 - len(r[2:32]))
            students.append({"sr": r[0] if r[0] is not None else k + 1, "name": name,
                             "marks": {co: [_cell_num(vals[ci * 6 + ti], f"{name} {co} {t}", errors)
                                            for ti, t in enumerate(TOOLS)] for ci, co in enumerate(COS)},
                             "end": _cell_num(ends[k][1], f"{name} End Term", errors)})
    else:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        if not rows or [str(h or "").strip() for h in rows[0][:33]] != template_header():
            return [], ["Unrecognised file: upload the original course workbook (with 'Tab 3') or the template from 'Download template'"]
        for k, r in enumerate(rows[1:]):
            r = list(r) + [None] * (33 - len(r))
            name = str(r[1] or "").strip()
            if not name:
                continue
            students.append({"sr": r[0] if r[0] is not None else k + 1, "name": name,
                             "marks": {co: [_cell_num(r[2 + ci * 6 + ti], f"{name} {co} {t}", errors)
                                            for ti, t in enumerate(TOOLS)] for ci, co in enumerate(COS)},
                             "end": _cell_num(r[32], f"{name} End Term", errors)})

    if not students and not errors:
        errors.append("No student rows found in the uploaded file")
    for s in students:
        if isinstance(s["sr"], float) and s["sr"].is_integer():
            s["sr"] = int(s["sr"])
    for s in students:
        for co in COS:
            for t, m, mx in zip(TOOLS, s["marks"][co], cfg["max_marks"][co]):
                if m < 0 or (mx and m > mx):
                    errors.append(f"{s['name']} {co} {t}: {m:g} is outside 0-{mx:g}")
        if not 0 <= s["end"] <= cfg["end_max"]:
            errors.append(f"{s['name']} End Term: {s['end']:g} is outside 0-{cfg['end_max']:g}")
    return students, errors


def template_header() -> list[str]:
    return ["Sr", "Name"] + [f"{co} {t}" for co in COS for t in TOOLS] + ["End Term"]


# ---------------------------------------------------------------- excel

def build_excel(cfg: dict, res: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    rows = [["Field", "Value"],
            ["Course", f"{cfg['course_name']} ({cfg['course_code']})"],
            ["Class", cfg["class_name"]], ["Semester", cfg["semester"]],
            ["Internal / End term split", f"{cfg['internal_pct']:g}% / {100 - cfg['internal_pct']:g}%"],
            ["Direct / Indirect split", f"{cfg['direct_pct']:g}% / {100 - cfg['direct_pct']:g}%"],
            ["End term max marks", cfg["end_max"]],
            ["Level 3 / 2 / 1 thresholds (%)", f">={cfg['level3']:g} / >={cfg['level2']:g} / >={cfg['level1']:g}"],
            [], ["CO", "Tool", "Max marks", "Weight (%)"]]
    rows += [[co, t, cfg["max_marks"][co][i], cfg["weights"][co][i]] for co in COS for i, t in enumerate(TOOLS)]
    for r in rows:
        ws.append(r)
    style_sheet(ws)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 44

    ws = wb.create_sheet("Student_Scores")
    ws.append(["Sr", "Name"] + [f"{co} internal" for co in COS] + ["End term", "End term (/100)"] + [f"{co} total" for co in COS])
    for s in res["students"]:
        ws.append([s["sr"], s["name"]] + [s["internal"][co] for co in COS] + [s["end"], s["end_conv"]] + [s["total"][co] for co in COS])
    style_sheet(ws)
    ws.column_dimensions["B"].width = 32
    ws.freeze_panes = "C2"

    ws = wb.create_sheet("CO_Attainment")
    ws.append(["Course Outcome", "Target (%)", "% Direct attainment", "% Indirect attainment",
               "Total actual CO attainment", "CO attainment level", "IA CO level", "Final CO attainment", "Status"])
    for c in res["cos"]:
        ws.append([c["co"], c["target"], c["direct"], c["indirect"], c["total"], c["level"], c["ia_level"], c["final"], c["status"]])
    ws.append(["Average CO Attainment", "", "", "", "", "", "", res["avg_final"], ""])
    style_sheet(ws)

    ws = wb.create_sheet("PO_Attainment")
    ws.append(["CO"] + POS)
    for co in COS:
        ws.append([co] + cfg["mapping"][co])
    ws.append(["Target attainment"] + [p["target"] for p in res["pos"]])
    ws.append(["Actual attainment"] + [p["actual"] for p in res["pos"]])
    style_sheet(ws)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- sqlite

SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL,
    course_name TEXT NOT NULL,
    course_code TEXT,
    inputs_json TEXT NOT NULL,
    avg_co_attainment REAL NOT NULL,
    student_count INTEGER NOT NULL,
    excel BLOB NOT NULL
);
CREATE TABLE IF NOT EXISTS student_scores (
    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    sr TEXT, name TEXT NOT NULL, end_term REAL, end_term_100 REAL,
    co1_internal REAL, co2_internal REAL, co3_internal REAL, co4_internal REAL, co5_internal REAL,
    co1_total REAL, co2_total REAL, co3_total REAL, co4_total REAL, co5_total REAL
);
CREATE TABLE IF NOT EXISTS co_attainment (
    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    co TEXT NOT NULL, target REAL, direct REAL, indirect REAL, total REAL,
    level INTEGER, ia_level REAL, final REAL, status TEXT
);
CREATE TABLE IF NOT EXISTS po_attainment (
    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    po TEXT NOT NULL, target REAL, actual REAL
);
"""


def db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
        g.db.executescript(SCHEMA)
    return g.db


def save_run(cfg: dict, res: dict, excel: bytes) -> int:
    con = db()
    with con:  # one transaction: all or nothing
        cur = con.execute(
            "INSERT INTO runs (created_at, course_name, course_code, inputs_json, avg_co_attainment, student_count, excel)"
            " VALUES (?, ?, ?, ?, ?, ?, ?)",
            (datetime.now().isoformat(timespec="seconds"), cfg["course_name"], cfg["course_code"],
             json.dumps(cfg), res["avg_final"], len(res["students"]), excel))
        run_id = cur.lastrowid
        con.executemany(
            "INSERT INTO student_scores VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            [(run_id, str(s["sr"]), s["name"], s["end"], s["end_conv"], *[s["internal"][co] for co in COS],
              *[s["total"][co] for co in COS]) for s in res["students"]])
        con.executemany("INSERT INTO co_attainment VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        [(run_id, c["co"], c["target"], c["direct"], c["indirect"], c["total"], c["level"],
                          c["ia_level"], c["final"], c["status"]) for c in res["cos"]])
        con.executemany("INSERT INTO po_attainment VALUES (?, ?, ?, ?)",
                        [(run_id, p["po"], p["target"], p["actual"]) for p in res["pos"]])
    return run_id


# ---------------------------------------------------------------- web

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024


@app.template_filter("fmt")
def fmt(x):
    return "" if x is None else f"{x:g}" if isinstance(x, (int, float)) else x


@app.teardown_appcontext
def close_db(_exc):
    con = g.pop("db", None)
    if con is not None:
        con.close()


def form_values(form=None) -> dict:
    """Flat name->value dict for the form: submitted values win, else DEFAULTS."""
    d = DEFAULTS
    v = {k: d[k] for k in ("course_name", "course_code", "class_name", "semester",
                           "level3", "level2", "level1", "internal_pct", "direct_pct", "end_max")}
    for co in COS:
        v.update({f"map_{co}_{po}": d["mapping"][co][i] for i, po in enumerate(POS)})
        v.update({f"max_{co}_{t}": d["max_marks"][co][i] for i, t in enumerate(TOOLS)})
        v.update({f"w_{co}_{t}": d["weights"][co][i] for i, t in enumerate(TOOLS)})
        for key in ("target", "indirect", "ia_level"):
            v[f"{key}_{co}"] = d[key][co]
    if form:
        v.update({k: form[k] for k in v if k in form})
    return v


def page(errors=None, form=None):
    runs = db().execute("SELECT id, created_at, course_name, course_code, avg_co_attainment, student_count"
                        " FROM runs ORDER BY id DESC LIMIT 50").fetchall()
    return render_template("index.html", v=form_values(form), errors=errors or [], runs=runs,
                           COS=COS, POS=POS, TOOLS=TOOLS, default_exists=DEFAULT_STUDENTS.exists())


@app.get("/")
def index():
    return page()


@app.post("/run")
def run():
    cfg, errors = parse_form(request.form)
    upload = request.files.get("students")
    if upload and upload.filename:
        data = upload.read()
    elif DEFAULT_STUDENTS.exists():
        data = DEFAULT_STUDENTS.read_bytes()
    else:
        errors.append("Upload a student marks file")
        data = None
    if errors:  # don't parse marks against a broken config
        return page(errors, request.form), 400
    students, errors = parse_students(data, cfg)
    if errors:
        more = len(errors) - 30
        return page(errors[:30] + ([f"... and {more} more"] if more > 0 else []), request.form), 400
    res = compute(cfg, students)
    run_id = save_run(cfg, res, build_excel(cfg, res))
    return redirect(url_for("result", run_id=run_id))


@app.get("/run/<int:run_id>")
def result(run_id: int):
    con = db()
    run_row = con.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
    if run_row is None:
        abort(404)
    return render_template(
        "result.html", run=run_row, cfg=json.loads(run_row["inputs_json"]), COS=COS, POS=POS,
        cos=con.execute("SELECT * FROM co_attainment WHERE run_id = ? ORDER BY co", (run_id,)).fetchall(),
        pos=con.execute("SELECT * FROM po_attainment WHERE run_id = ? ORDER BY rowid", (run_id,)).fetchall(),
        students=con.execute("SELECT * FROM student_scores WHERE run_id = ? ORDER BY rowid", (run_id,)).fetchall())


@app.get("/run/<int:run_id>/excel")
def download(run_id: int):
    row = db().execute("SELECT excel, course_code FROM runs WHERE id = ?", (run_id,)).fetchone()
    if row is None:
        abort(404)
    code = re.sub(r"[^A-Za-z0-9_-]", "", row["course_code"] or "") or "course"
    return send_file(BytesIO(row["excel"]), as_attachment=True,
                     download_name=f"course_attainment_{code}_run{run_id}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/run/<int:run_id>/delete")
def delete(run_id: int):
    con = db()
    with con:
        con.execute("DELETE FROM runs WHERE id = ?", (run_id,))
    return redirect(url_for("index"))


@app.get("/template")
def template():
    wb = Workbook()
    wb.active.append(template_header())
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="student_marks_template.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(debug=False)

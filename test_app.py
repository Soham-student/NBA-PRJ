"""Run: python test_app.py"""
import tempfile
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

import app as A


def test_compute_by_hand():
    cfg = dict(A.DEFAULTS)
    s = {"sr": 1, "name": "X", "end": 20, "marks": {co: [5, 10, 15, 0, 0, 30] for co in A.COS}}
    s["marks"]["CO1"] = [3, 0, 6, 0, 0, 12]
    res = A.compute(cfg, [s])
    st = res["students"][0]
    # CO1 internal = 3*17.14/5 + 6*40/15 + 12*42.86/30 = 10.28 + 16 + 17.14
    assert st["internal"]["CO1"] == 43.42, st
    # total = round(43.42*0.6) + round(50*0.4) = 26.05 + 20
    assert st["total"]["CO1"] == 46.05, st
    assert st["total"]["CO3"] == 80.0  # full marks: 100*0.6 + 50*0.4
    co1 = res["cos"][0]
    assert co1["total"] == A.r2(0.8 * 46.05 + 0.2 * 80) == 52.84
    assert co1["level"] == 2 and co1["status"] == "Not attained"  # target 60
    assert co1["final"] == 2.2  # 0.8*2 + 0.2*3
    po4 = res["pos"][3]
    assert po4["target"] == 0 and po4["actual"] == 0  # unmapped PO
    assert A.r2(2.675) == 2.68 and A.r2(0.125) == 0.13  # half-up like Excel


def test_web_end_to_end():
    with tempfile.TemporaryDirectory() as d:
        A.DB_PATH = Path(d) / "t.db"
        c = A.app.test_client()
        assert c.get("/").status_code == 200

        form = {k: str(v) for k, v in A.form_values().items()}
        r = c.post("/run", data=form)  # no upload -> repo input sheet
        assert r.status_code == 302, r.data[:2000]
        page = c.get(r.headers["Location"])
        assert page.status_code == 200 and b"Prajwal Gadkar" in page.data and b"Table F" in page.data

        x = c.get(r.headers["Location"] + "/excel")
        wb = load_workbook(BytesIO(x.data))
        assert wb.sheetnames == ["Inputs", "Student_Scores", "CO_Attainment", "PO_Attainment"]
        assert wb["Student_Scores"].max_row == 292  # header + 291 students

        con = A.sqlite3.connect(A.DB_PATH)
        assert con.execute("SELECT COUNT(*) FROM student_scores").fetchone()[0] == 291
        assert con.execute("SELECT COUNT(*) FROM co_attainment").fetchone()[0] == 5

        # bad input is rejected with a message, nothing saved
        bad = dict(form, w_CO1_LCA="10")
        r = c.post("/run", data=bad)
        assert r.status_code == 400 and b"must be 100" in r.data
        assert con.execute("SELECT COUNT(*) FROM runs").fetchone()[0] == 1

        # template upload path
        wb = Workbook()
        wb.active.append(A.template_header())
        wb.active.append([1, "Test Student"] + [1] * 30 + [40])
        buf = BytesIO()
        wb.save(buf)
        r = c.post("/run", data=dict(form, students=(BytesIO(buf.getvalue()), "m.xlsx")),
                   content_type="multipart/form-data")
        assert r.status_code == 302, r.data[:2000]

        # delete cascades
        c.post("/run/1/delete")
        assert con.execute("SELECT COUNT(*) FROM student_scores WHERE run_id = 1").fetchone()[0] == 0
        con.close()


if __name__ == "__main__":
    test_compute_by_hand()
    test_web_end_to_end()
    print("all tests passed")

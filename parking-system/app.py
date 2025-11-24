import os
import re
import sqlite3
from datetime import datetime
from io import BytesIO

from flask import (
    Flask, render_template, request, redirect,
    url_for, flash, send_file, session
)

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = "parking_secret_key"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "parking.db")

LOGIN_PASSWORD = "1112"


# -------------------------------- 로그인 --------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        pw = request.form.get("password")
        if pw == LOGIN_PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("index"))
        flash("비밀번호가 틀렸습니다.")
    return render_template("login.html")


# -------------------------------- 메인 페이지 --------------------------------

@app.route("/", methods=["GET"])
def index():
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    view_date = request.args.get("view_date") or datetime.now().strftime("%Y-%m-%d")
    query_plate = request.args.get("q", "")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # 전체 차량 목록(자동완성용)
    cur.execute("SELECT DISTINCT plate FROM records ORDER BY plate")
    all_plates = [row[0] for row in cur.fetchall()]

    # 선택 날짜 차량 목록
    cur.execute("""
        SELECT
            id, plate,
            (SELECT COUNT(*) FROM records r2 WHERE r2.plate = r.plate) AS total_cnt
        FROM records r
        WHERE date = ?
        ORDER BY plate
    """, (view_date,))
    day_rows = cur.fetchall()

    # 조회 기능
    search_result = None
    if query_plate:
        cur.execute("""
            SELECT
                plate,
                COUNT(*) AS total_cnt,
                MIN(date) AS first_date,
                MAX(date) AS last_date
            FROM records
            WHERE plate = ?
        """, (query_plate,))
        row = cur.fetchone()
        if row and row[1] > 0:
            search_result = {
                "plate": row[0],
                "total_cnt": row[1],
                "first_date": row[2],
                "last_date": row[3]
            }
        else:
            search_result = {"plate": query_plate, "total_cnt": 0}

    conn.close()

    return render_template(
        "index.html",
        today=datetime.now().strftime("%Y-%m-%d"),
        view_date=view_date,
        all_plates=all_plates,
        day_rows=day_rows,
        day_count=len(day_rows),
        query_plate=query_plate,
        search_result=search_result
    )


# -------------------------------- 차량 등록 --------------------------------

@app.route("/add", methods=["POST"])
def add():
    date = request.form.get("date")
    plates = request.form.getlist("plate")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    for p in plates:
        if not p.strip():
            continue
        cur.execute("INSERT INTO records (date, plate) VALUES (?, ?)", (date, p.strip()))
    conn.commit()
    conn.close()

    flash("등록되었습니다.")
    return redirect(url_for("index", view_date=date))


# -------------------------------- 삭제 --------------------------------

@app.route("/delete", methods=["POST"])
def delete():
    rid = request.form.get("id")
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("DELETE FROM records WHERE id = ?", (rid,))
    conn.commit()
    conn.close()
    flash("삭제 완료")
    return redirect(url_for("index"))


# -------------------------------- 엑셀 EXPORT (openpyxl) --------------------------------

@app.route("/export")
def export():
    scope = request.args.get("scope", "day")
    view_date = request.args.get("date")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # summary
    cur.execute("""
        SELECT
            plate,
            COUNT(*) AS total_cnt,
            MIN(date) AS first_date,
            MAX(date) AS last_date
        FROM records
        GROUP BY plate
        ORDER BY plate
    """)
    summary_rows = cur.fetchall()

    # 전체
    cur.execute("SELECT date, plate FROM records ORDER BY date, plate")
    all_rows = cur.fetchall()
    conn.close()

    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"

    ws_summary.append(["차량번호", "누적횟수", "최초입차일", "최근입차일"])
    for row in summary_rows:
        ws_summary.append(row)

    grouped = {}
    for date, plate in all_rows:
        grouped.setdefault(date, []).append(plate)

    for date, plates in grouped.items():

        safe_date = date.replace("/", "-")[:31]

        ws = wb.create_sheet(title=safe_date)
        ws.append(["차량번호", "누적횟수", "최초입차일", "최근입차일"])

        for p in plates:
            match = next((r for r in summary_rows if r[0] == p), None)
            ws.append(match if match else [p, 1, date, date])

    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))
    center = Alignment(horizontal="center", vertical="center")
    bold = Font(bold=True)

    for sheet in wb.worksheets:
        for c in sheet[1]:
            c.font = bold

        for row in sheet.rows:
            for cell in row:
                cell.alignment = center
                cell.border = border

        for col in range(1, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 15

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)

    filename = datetime.now().strftime("%Y-%m-%d") + ".xlsx"

    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# -------------------------------- 로그아웃 --------------------------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


if __name__ == "__main__":
    app.run(debug=True)

"""Xuat bao cao ket qua: Excel (results.xlsx) va HTML (report.html)."""
import html
import os
from collections import Counter

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FONT = "Times New Roman"

FILLS = {
    "Pass": PatternFill("solid", fgColor="D4EDDA"),
    "Fail": PatternFill("solid", fgColor="F8D7DA"),
    "Blocked": PatternFill("solid", fgColor="FFF3CD"),
    "Error": PatternFill("solid", fgColor="F5C6CB"),
}

STATUS_FONTS = {
    "Pass": Font(name=FONT, bold=True, size=10, color="155724"),
    "Fail": Font(name=FONT, bold=True, size=10, color="721C24"),
    "Blocked": Font(name=FONT, bold=True, size=10, color="856404"),
    "Error": Font(name=FONT, bold=True, size=10, color="721C24"),
}

THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

STATUSES = ["Pass", "Fail", "Blocked", "Error"]


def load_cases(testdata_path: str, sheet: str) -> list:
    """Doc cac test case cua mot sheet thanh list dict (khoa = ten cot viet thuong, '_' thay khoang trang)."""
    wb = load_workbook(testdata_path)
    ws = wb[sheet]
    headers = [str(c.value or "").strip().lower().replace(" ", "_") for c in ws[1]]
    cases = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        case = dict(zip(headers, row))
        case["id"] = str(case["test_case_id"]).strip()
        cases.append(case)
    return cases


def write_excel(testdata_path: str, out_path: str, results: list, meta: dict) -> None:
    wb = load_workbook(testdata_path)
    by_sheet = {}
    for r in results:
        by_sheet.setdefault(r.sheet, {})[r.id] = r

    ranges = {}
    for sheet, by_id in by_sheet.items():
        ws = wb[sheet]
        header = {str(c.value or "").strip(): c.column for c in ws[1]}
        id_col, st_col, ac_col = header["Test Case ID"], header["Status"], header["Actual Result"]

        ws.row_dimensions[1].height = 28
        for col_idx in range(1, ws.max_column + 1):
            h_cell = ws.cell(1, col_idx)
            h_cell.font = Font(name=FONT, bold=True, size=11, color="FFFFFF")
            h_cell.fill = PatternFill("solid", fgColor="1F4E78")
            h_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            h_cell.border = THIN_BORDER

        for row in range(2, ws.max_row + 1):
            ws.row_dimensions[row].height = 34
            row_bg = "F8FAFC" if row % 2 == 0 else "FFFFFF"

            for col_idx in range(1, ws.max_column + 1):
                c = ws.cell(row, col_idx)
                c.font = Font(name=FONT, size=10)
                c.fill = PatternFill("solid", fgColor=row_bg)
                c.alignment = Alignment(wrap_text=True, vertical="center")
                c.border = THIN_BORDER

            tc_id = str(ws.cell(row, id_col).value or "").strip()
            if tc_id in by_id:
                res = by_id[tc_id]
                st = ws.cell(row, st_col, res.status)
                st.font = STATUS_FONTS.get(res.status, Font(name=FONT, bold=True, size=10))
                st.alignment = Alignment(horizontal="center", vertical="center")
                if res.status in FILLS:
                    st.fill = FILLS[res.status]

                ac_cell = ws.cell(row, ac_col, res.actual)
                ac_cell.alignment = Alignment(wrap_text=True, vertical="center")

        col = ws.cell(1, st_col).column_letter
        ranges[sheet] = f"'{sheet}'!${col}$2:${col}${ws.max_row}"

    sm = wb.create_sheet("Summary", 0)
    sm.views.sheetView[0].showGridLines = True

    sm.merge_cells("A1:H1")
    sm["A1"] = "BÁO CÁO TỔNG HỢP KẾT QUẢ KIỂM THỬ TỰ ĐỘNG"
    sm["A1"].font = Font(name=FONT, size=15, bold=True, color="1F4E78")
    sm["A1"].alignment = Alignment(horizontal="center", vertical="center")
    sm.row_dimensions[1].height = 35

    meta_rows = [
        ("Thời gian chạy:", meta.get("time", "")),
        ("Website kiểm thử:", meta.get("base_url", "")),
        ("Trình duyệt:", meta.get("browser", "")),
    ]
    for r, (k, v) in enumerate(meta_rows, start=2):
        sm.row_dimensions[r].height = 22
        c_key = sm.cell(r, 1, k)
        c_key.font = Font(name=FONT, bold=True, size=10, color="333333")
        c_val = sm.cell(r, 2, v)
        c_val.font = Font(name=FONT, size=10, color="0056B3")

    sm.row_dimensions[6].height = 26
    head = ["Chức năng", "Tổng", "Pass", "Fail", "Blocked", "Error", "Not Run", "Tỷ lệ Pass"]
    for c, h in enumerate(head, start=1):
        cell = sm.cell(6, c, h)
        cell.font = Font(name=FONT, bold=True, size=11, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    r = 7
    for sheet, rng in ranges.items():
        sm.row_dimensions[r].height = 24
        sm.cell(r, 1, sheet).alignment = Alignment(horizontal="left", vertical="center")
        sm.cell(r, 2, f"=COUNTA({rng})").alignment = Alignment(horizontal="center", vertical="center")
        for c, s in enumerate(["Pass", "Fail", "Blocked", "Error", "Not Run"], start=3):
            cell_s = sm.cell(r, c, f'=COUNTIF({rng},"{s}")')
            cell_s.alignment = Alignment(horizontal="center", vertical="center")
        rate_cell = sm.cell(r, 8, f"=IFERROR(C{r}/B{r},0)")
        rate_cell.number_format = "0.0%"
        rate_cell.alignment = Alignment(horizontal="center", vertical="center")
        r += 1

    if ranges:
        sm.row_dimensions[r].height = 26
        sm.cell(r, 1, "Tổng cộng").alignment = Alignment(horizontal="left", vertical="center")
        for c in range(2, 8):
            col = sm.cell(6, c).column_letter
            c_sum = sm.cell(r, c, f"=SUM({col}7:{col}{r - 1})")
            c_sum.alignment = Alignment(horizontal="center", vertical="center")
        c_total_rate = sm.cell(r, 8, f"=IFERROR(C{r}/B{r},0)")
        c_total_rate.number_format = "0.0%"
        c_total_rate.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx in range(7, r + 1):
        is_total = (row_idx == r)
        for col_idx in range(1, 9):
            cell = sm.cell(row_idx, col_idx)
            cell.font = Font(name=FONT, size=10, bold=is_total)
            cell.border = THIN_BORDER
            if is_total:
                cell.fill = PatternFill("solid", fgColor="D9E1F2")

    col_widths = {"A": 22, "B": 36, "C": 10, "D": 10, "E": 12, "F": 10, "G": 12, "H": 14}
    for col_letter, width in col_widths.items():
        sm.column_dimensions[col_letter].width = width

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)


def write_html(out_path: str, results: list, meta: dict) -> None:
    count = Counter(r.status for r in results)
    total = len(results)
    pass_pct = (count["Pass"] * 100 / total) if total else 0
    fail_pct = (count["Fail"] * 100 / total) if total else 0
    block_pct = (count["Blocked"] * 100 / total) if total else 0
    err_pct = (count["Error"] * 100 / total) if total else 0
    rate = f"{pass_pct:.1f}%"

    rows = []
    for r in results:
        shot = (
            f'<a class="btn-shot" href="{html.escape(r.screenshot)}" target="_blank">Xem ảnh</a>'
            if r.screenshot
            else '<span class="no-shot">—</span>'
        )
        rows.append(
            f'<tr data-status="{html.escape(r.status)}">'
            f'<td><span class="tag-sheet">{html.escape(r.sheet)}</span></td>'
            f'<td class="tc-id">{html.escape(r.id)}</td>'
            f'<td class="tc-title">{html.escape(r.title)}</td>'
            f'<td>{html.escape(r.expected)}</td>'
            f'<td style="text-align:center"><span class="badge {r.status}">{r.status}</span></td>'
            f'<td class="tc-actual">{html.escape(r.actual)}</td>'
            f'<td style="text-align:center">{shot}</td>'
            f'</tr>'
        )

    page = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Báo cáo Kiểm thử Tự động</title>
<style>
  :root {{
    --primary: #1e3a8a;
    --primary-light: #2563eb;
    --bg: #f1f5f9;
    --card-bg: #ffffff;
    --text: #1e293b;
    --muted: #64748b;
    --border: #e2e8f0;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Segoe UI', Roboto, Arial, sans-serif;
    background-color: var(--bg);
    color: var(--text);
    line-height: 1.5;
    padding-bottom: 40px;
  }}
  .header {{
    background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
    color: #ffffff;
    padding: 28px 36px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
  }}
  .header-inner {{
    max-width: 1360px;
    margin: 0 auto;
  }}
  .header h1 {{
    font-size: 24px;
    font-weight: 700;
    letter-spacing: 0.3px;
    margin-bottom: 10px;
  }}
  .meta-pills {{
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    font-size: 13px;
  }}
  .meta-pill {{
    background: rgba(255, 255, 255, 0.12);
    padding: 5px 12px;
    border-radius: 20px;
    border: 1px solid rgba(255, 255, 255, 0.2);
  }}
  .container {{
    max-width: 1360px;
    margin: 24px auto 0;
    padding: 0 24px;
  }}
  .cards {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
    gap: 16px;
    margin-bottom: 20px;
  }}
  .card {{
    background: var(--card-bg);
    border-radius: 12px;
    padding: 18px 20px;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.06);
    border-top: 4px solid #94a3b8;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
  }}
  .card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 6px 14px rgba(15, 23, 42, 0.1);
  }}
  .card .label {{
    font-size: 13px;
    font-weight: 600;
    color: var(--muted);
    text-transform: uppercase;
  }}
  .card .num {{
    font-size: 30px;
    font-weight: 800;
    margin-top: 4px;
  }}
  .card.c-total {{ border-top-color: #2563eb; }}
  .card.c-total .num {{ color: #1e3a8a; }}
  .card.c-pass {{ border-top-color: #16a34a; }}
  .card.c-pass .num {{ color: #15803d; }}
  .card.c-fail {{ border-top-color: #dc2626; }}
  .card.c-fail .num {{ color: #b91c1c; }}
  .card.c-blocked {{ border-top-color: #d97706; }}
  .card.c-blocked .num {{ color: #b45309; }}
  .card.c-error {{ border-top-color: #ea580c; }}
  .card.c-error .num {{ color: #c2410c; }}
  .card.c-rate {{ border-top-color: #0d9488; }}
  .card.c-rate .num {{ color: #0f766e; }}

  .progress-box {{
    background: var(--card-bg);
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 20px;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.06);
  }}
  .progress-header {{
    display: flex;
    justify-content: space-between;
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 8px;
    color: var(--muted);
  }}
  .progress-bar {{
    height: 12px;
    border-radius: 999px;
    background: #e2e8f0;
    overflow: hidden;
    display: flex;
  }}
  .pb-pass {{ background: #22c55e; width: {pass_pct:.2f}%; }}
  .pb-fail {{ background: #ef4444; width: {fail_pct:.2f}%; }}
  .pb-blocked {{ background: #f59e0b; width: {block_pct:.2f}%; }}
  .pb-error {{ background: #f97316; width: {err_pct:.2f}%; }}

  .toolbar {{
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    background: var(--card-bg);
    padding: 14px 20px;
    border-radius: 12px 12px 0 0;
    border-bottom: 1px solid var(--border);
  }}
  .filters {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
  }}
  .filter-btn {{
    border: 1px solid var(--border);
    background: #f8fafc;
    color: var(--text);
    padding: 6px 14px;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.15s;
  }}
  .filter-btn:hover {{ background: #e2e8f0; }}
  .filter-btn.active {{
    background: var(--primary);
    color: #ffffff;
    border-color: var(--primary);
  }}
  .search-input {{
    padding: 7px 14px;
    border: 1px solid var(--border);
    border-radius: 8px;
    font-size: 13px;
    min-width: 260px;
    outline: none;
  }}
  .search-input:focus {{
    border-color: var(--primary-light);
    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
  }}

  .table-wrap {{
    background: var(--card-bg);
    border-radius: 0 0 12px 12px;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.06);
    overflow-x: auto;
  }}
  table {{
    border-collapse: collapse;
    width: 100%;
    font-size: 13.5px;
  }}
  th {{
    background: #f8fafc;
    color: #334155;
    font-weight: 700;
    text-transform: uppercase;
    font-size: 12px;
    letter-spacing: 0.4px;
    padding: 12px 14px;
    border-bottom: 2px solid var(--border);
    text-align: left;
  }}
  td {{
    padding: 12px 14px;
    border-bottom: 1px solid var(--border);
    vertical-align: top;
  }}
  tbody tr:hover {{
    background-color: #f8fafc;
  }}
  .tag-sheet {{
    display: inline-block;
    background: #e0e7ff;
    color: #1e40af;
    font-weight: 600;
    font-size: 12px;
    padding: 2px 8px;
    border-radius: 6px;
  }}
  .tc-id {{
    font-family: Consolas, monospace;
    font-weight: 700;
    color: #0f172a;
    white-space: nowrap;
  }}
  .tc-title {{ font-weight: 600; }}
  .tc-actual {{ color: #334155; }}
  .badge {{
    display: inline-block;
    padding: 4px 12px;
    border-radius: 999px;
    font-weight: 700;
    font-size: 12px;
  }}
  .badge.Pass {{ background: #dcfce7; color: #166534; }}
  .badge.Fail {{ background: #fee2e2; color: #991b1b; }}
  .badge.Blocked {{ background: #fef3c7; color: #92400e; }}
  .badge.Error {{ background: #ffedd5; color: #9a3412; }}

  .btn-shot {{
    display: inline-block;
    background: #eff6ff;
    color: #1d4ed8;
    border: 1px solid #bfdbfe;
    padding: 4px 10px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    text-decoration: none;
    white-space: nowrap;
  }}
  .btn-shot:hover {{ background: #dbeafe; }}
  .no-shot {{ color: #94a3b8; }}
</style>
</head>
<body>
  <div class="header">
    <div class="header-inner">
      <h1>BÁO CÁO KIỂM THỬ TỰ ĐỘNG: ĐĂNG KÝ &amp; ĐĂNG NHẬP</h1>
      <div class="meta-pills">
        <span class="meta-pill">Thời gian: <b>{html.escape(meta.get('time', ''))}</b></span>
        <span class="meta-pill">Website: <b>{html.escape(meta.get('base_url', ''))}</b></span>
        <span class="meta-pill">Trình duyệt: <b>{html.escape(meta.get('browser', ''))}</b></span>
      </div>
    </div>
  </div>

  <div class="container">
    <div class="cards">
      <div class="card c-total"><div class="label">Tổng Test Case</div><div class="num">{total}</div></div>
      <div class="card c-pass"><div class="label">Pass</div><div class="num">{count['Pass']}</div></div>
      <div class="card c-fail"><div class="label">Fail</div><div class="num">{count['Fail']}</div></div>
      <div class="card c-blocked"><div class="label">Blocked</div><div class="num">{count['Blocked']}</div></div>
      <div class="card c-error"><div class="label">Error</div><div class="num">{count['Error']}</div></div>
      <div class="card c-rate"><div class="label">Tỷ lệ Pass</div><div class="num">{rate}</div></div>
    </div>

    <div class="progress-box">
      <div class="progress-header">
        <span>Phân bổ trạng thái thực thi</span>
        <span>Pass: {pass_pct:.1f}% | Fail: {fail_pct:.1f}% | Blocked: {block_pct:.1f}% | Error: {err_pct:.1f}%</span>
      </div>
      <div class="progress-bar">
        <div class="pb-pass" title="Pass: {count['Pass']}"></div>
        <div class="pb-fail" title="Fail: {count['Fail']}"></div>
        <div class="pb-blocked" title="Blocked: {count['Blocked']}"></div>
        <div class="pb-error" title="Error: {count['Error']}"></div>
      </div>
    </div>

    <div class="toolbar">
      <div class="filters">
        <button class="filter-btn active" onclick="filterTable('ALL', this)">Tất cả ({total})</button>
        <button class="filter-btn" onclick="filterTable('Pass', this)">Pass ({count['Pass']})</button>
        <button class="filter-btn" onclick="filterTable('Fail', this)">Fail ({count['Fail']})</button>
        <button class="filter-btn" onclick="filterTable('Blocked', this)">Blocked ({count['Blocked']})</button>
        <button class="filter-btn" onclick="filterTable('Error', this)">Error ({count['Error']})</button>
      </div>
      <input type="text" id="searchInput" class="search-input" onkeyup="searchTable()" placeholder="Tìm theo Mã TC, mô tả, kết quả...">
    </div>

    <div class="table-wrap">
      <table id="reportTable">
        <thead>
          <tr>
            <th>Chức năng</th>
            <th>Mã TC</th>
            <th>Mô tả kịch bản</th>
            <th>Kết quả kỳ vọng</th>
            <th style="text-align:center">Trạng thái</th>
            <th>Kết quả thực tế</th>
            <th style="text-align:center">Bằng chứng</th>
          </tr>
        </thead>
        <tbody>
          {''.join(rows)}
        </tbody>
      </table>
    </div>
  </div>

<script>
  let currentStatus = 'ALL';
  function filterTable(status, btn) {{
    currentStatus = status;
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    applyFilters();
  }}
  function searchTable() {{
    applyFilters();
  }}
  function applyFilters() {{
    const keyword = document.getElementById('searchInput').value.toLowerCase();
    const trs = document.querySelectorAll('#reportTable tbody tr');
    trs.forEach(tr => {{
      const st = tr.getAttribute('data-status');
      const text = tr.innerText.toLowerCase();
      const matchStatus = (currentStatus === 'ALL' || st === currentStatus);
      const matchSearch = text.includes(keyword);
      tr.style.display = (matchStatus && matchSearch) ? '' : 'none';
    }});
  }}
</script>
</body>
</html>"""

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(page)
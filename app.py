"""Delivery Transit Analysis (D+0): 1) generate reports from raw CSVs, 2) compare two periods."""
import io, re, hmac, datetime, warnings
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter as L
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

NAME = "Name of the Division"
DRR = "D+0 (Delivered+Returned+Redirected)"
HEAD = ["Received", "Invoiced", "Invoiced %", "D+0 Delivered", "D+0 Delivered %",
        "D+0 Redirected", "D+0 Redirected %", "D+0 Returned", "D+0 Returned %",
        DRR, DRR + " (Received) %", DRR + " (Invoiced) %"]
PCT = {2, 4, 6, 8, 10, 11}              # positions in HEAD that are percentages
COMP_HEAD = ["Received Count Difference", "Invoiced % Increase or Decrease",
             "D+0 Delivered % Increase or Decrease", "D+0 Redirected % Increase or Decrease",
             "D+0 Returned % Increase or Decrease",
             DRR + " (Received) % Increase or Decrease", DRR + " (Invoiced) % Increase or Decrease"]
FMT_HEAD = ["0.00%" if i in PCT else "#,##0" for i in range(12)]
FMT_COMP = ["+#,##0;-#,##0;0"] + ["+0.00%;-0.00%;0.00%"] * 6
PERIOD = re.compile(r"\d{1,2}\.\d{1,2}\.\d{4}\s*to\s*\d{1,2}\.\d{1,2}\.\d{4}", re.I)
RAW_COLS = ["Region Name", "Received", "Same Day Invoiced", "D0 Delivered", "D0 Redirected", "D0 Returned"]
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GREEN, YELLOW = PatternFill("solid", fgColor="92D050"), PatternFill("solid", fgColor="FFFF00")
SIDE = Side(style="thin")
BOX = Border(left=SIDE, right=SIDE, top=SIDE, bottom=SIDE)


def to_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(re.sub(r"[,\s%]", "", str(v)))
    except ValueError:
        return 0.0


def sd(a, b):
    return a / b if b else 0.0


def norm(x):
    return re.sub(r"\s+", "", str(x)).lower()


def metrics(c):
    rec, inv, dl, rd, rt = c
    drr = dl + rd + rt
    return [rec, inv, sd(inv, rec), dl, sd(dl, rec), rd, sd(rd, rec), rt, sd(rt, rec),
            drr, sd(drr, rec), sd(drr, inv)]


# ---------- Excel writing ----------
def write_block(ws, col0, title, heads, table, fmts, r0=1):
    """One table: title row at r0, headers at r0+1, data below. Last table row = Total (yellow)."""
    ws.merge_cells(start_row=r0, start_column=col0, end_row=r0, end_column=col0 + len(heads) - 1)
    c = ws.cell(r0, col0, title)
    c.fill, c.font = GREEN, Font(name="Calibri", bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center")
    for j, h in enumerate(heads):
        c = ws.cell(r0 + 1, col0 + j, h)
        c.font, c.border = Font(name="Calibri", bold=True), BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for k, row in enumerate(table):
        last = k == len(table) - 1
        for j, v in enumerate(row):
            c = ws.cell(r0 + 2 + k, col0 + j, v)
            c.border, c.font = BOX, Font(name="Calibri", bold=last)
            if last:
                c.fill = YELLOW
            if fmts[j]:
                c.number_format = fmts[j]
    ws.row_dimensions[r0 + 1].height = 80


def report_table(rows):
    return [[i + 1 if r[0] != "Total" else "", r[0]] + r[1] for i, r in enumerate(rows)]


def build_report_sheet(ws, blocks):
    """blocks = [(title, rows)] stacked one below the other."""
    r0 = 1
    for title, rows in blocks:
        write_block(ws, 1, title, ["S.No.", NAME] + HEAD, report_table(rows), [None, None] + FMT_HEAD, r0)
        r0 += len(rows) + 3
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 7, 30
    for col in range(3, 15):
        ws.column_dimensions[L(col)].width = 14
    ws.freeze_panes = "C1"


def build_comparison_sheet(ws, blocks):
    """blocks = [(title1, title2, title_comp, rows)] stacked one below the other."""
    r0 = 1
    for t1, t2, tc, rows in blocks:
        write_block(ws, 1, t1, [NAME] + HEAD, [[r[0]] + r[1] for r in rows], [None] + FMT_HEAD, r0)
        write_block(ws, 15, t2, [NAME] + HEAD, [[r[0]] + r[2] for r in rows], [None] + FMT_HEAD, r0)
        write_block(ws, 29, tc, [NAME] + COMP_HEAD, [[r[0]] + r[3] for r in rows], [None] + FMT_COMP, r0)
        r0 += len(rows) + 3
    for col in range(1, 37):
        ws.column_dimensions[L(col)].width = 26 if col in (1, 15, 29) else 13
    ws.column_dimensions["N"].width = ws.column_dimensions["AB"].width = 3
    ws.freeze_panes = "B1"


def to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------- 1) Generate reports from raw CSVs ----------
def read_raw(f):
    """Uploaded raw file -> {region: [Received, Invoiced, Delivered, Redirected, Returned]} (file order)."""
    data = io.BytesIO(f.getvalue())
    df = pd.read_csv(data) if f.name.lower().endswith(".csv") else pd.read_excel(data)
    miss = [c for c in RAW_COLS if c not in df.columns]
    if miss:
        raise ValueError(f"{f.name}: these columns are missing: " + ", ".join(miss))
    df["Region Name"] = df["Region Name"].astype(str).str.split().str.join(" ")
    for c in RAW_COLS[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    g = df.groupby("Region Name", sort=False)[RAW_COLS[1:]].sum()
    return {n: [float(x) for x in v] for n, v in g.iterrows()}


def add_dicts(ds):
    out = {}
    for d in ds:
        for k, v in d.items():
            out[k] = [a + b for a, b in zip(out.get(k, [0.0] * 5), v)]
    return out


def sub_dicts(a, b):
    out = {k: [x - y for x, y in zip(v, b.get(k, [0.0] * 5))] for k, v in a.items()}
    for k, v in b.items():
        if k not in a:
            out[k] = [-x for x in v]
    return out


def to_rows(d):
    tot = [sum(c[i] for c in d.values()) for i in range(5)]
    return [(n, metrics(c)) for n, c in d.items()] + [("Total", metrics(tot))]


def make_reports(jobs, d1, d2):
    """jobs = [(sheet_name, report_name, [all-offices dicts], [BO dicts])]."""
    wb = Workbook()
    wb.remove(wb.active)
    previews = {}
    for sheet, rname, all_d, bo_d in jobs:
        a = add_dicts(all_d)
        scopes = [("All Offices", a)]
        if bo_d:
            b = add_dicts(bo_d)
            scopes += [("BOs", b), ("Excluding BOs", sub_dicts(a, b))]
        blocks = [(f"Delivery Transit Analysis {rname}- {s} -{d1:%d.%m.%Y} to {d2:%d.%m.%Y}", to_rows(d))
                  for s, d in scopes]
        build_report_sheet(wb.create_sheet(sheet), blocks)
        previews[sheet] = [(t, pd.DataFrame(report_table(r), columns=["S.No.", NAME] + HEAD)) for t, r in blocks]
    return to_bytes(wb), previews


# ---------- 1b) NDC / DC office-wise reports ----------
OFFICE_HEAD = ["S.No.", "Name of the Office", "Division"]


def read_offices(f):
    """Raw NDC/DC file -> ([(office, division, [5 counts])] in file order, set of Office Types)."""
    data = io.BytesIO(f.getvalue())
    df = pd.read_csv(data) if f.name.lower().endswith(".csv") else pd.read_excel(data)
    miss = [c for c in ["Office Name"] + RAW_COLS[1:] if c not in df.columns]
    if miss:
        raise ValueError(f"{f.name}: these columns are missing: " + ", ".join(miss))
    types = {str(t).strip().upper() for t in df["Office Type"].dropna()} if "Office Type" in df.columns else set()
    if "Division Name" not in df.columns:
        df["Division Name"] = ""
    df["Office Name"] = df["Office Name"].astype(str).str.split().str.join(" ")
    df["Division Name"] = df["Division Name"].fillna("").astype(str).str.split().str.join(" ")
    for c in RAW_COLS[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    g = df.groupby("Office Name", sort=False).agg({"Division Name": "first", **{c: "sum" for c in RAW_COLS[1:]}})
    return [(n, r["Division Name"], [float(r[c]) for c in RAW_COLS[1:]]) for n, r in g.iterrows()], types


def make_office_reports(jobs, d1, d2):
    """jobs = [(sheet_name, label, offices)] -> (workbook bytes, previews)."""
    wb = Workbook()
    wb.remove(wb.active)
    previews = {}
    for sheet, label, offices in jobs:
        title = f"Delivery Transit Analysis {label} -{d1:%d.%m.%Y} to {d2:%d.%m.%Y}"
        tot = [sum(o[2][i] for o in offices) for i in range(5)]
        table = ([[i + 1, n, dv] + metrics(c) for i, (n, dv, c) in enumerate(offices)]
                 + [["", "Total", ""] + metrics(tot)])
        ws = wb.create_sheet(sheet)
        write_block(ws, 1, title, OFFICE_HEAD + HEAD, table, [None] * 3 + FMT_HEAD)
        ws.column_dimensions["A"].width = 7
        ws.column_dimensions["B"].width = ws.column_dimensions["C"].width = 30
        for col in range(4, 16):
            ws.column_dimensions[L(col)].width = 14
        ws.freeze_panes = "C1"
        previews[sheet] = [(title, pd.DataFrame(table, columns=OFFICE_HEAD + HEAD))]
    return to_bytes(wb), previews


# ---------- 1c) Day-wise comparison (DCs / NDCs picked out of full all-offices files) ----------
GROUPS = {"DCs (DC / IDC)": re.compile(r"\bI?DC\b", re.I), "NDCs": re.compile(r"\bNDC\b", re.I),
          "NDCs + DCs": re.compile(r"\b(?:NDC|I?DC)\b", re.I)}
TOTAL_LABEL = "HQR"
BLUE, CREAM = PatternFill("solid", fgColor="DDEBF7"), PatternFill("solid", fgColor="FFF2CC")
GOOD, BAD = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFC7CE")


def pick_group(offices, rx):
    return {" ".join(n.lower().split()): (n, dv, c) for n, dv, c in offices if rx.search(n)}


def day_compare(g1, g2):
    """-> (rows sorted by % Inc/Dec, total). row = (office, division, counts_day1, counts_day2)."""
    zero = (None, "", [0.0] * 5)
    rows = []
    for k in list(g2) + [k for k in g1 if k not in g2]:
        a, b = g1.get(k, zero), g2.get(k, zero)
        div = re.sub(r"\s*division\s*$", "", b[1] or a[1], flags=re.I)
        rows.append((b[0] or a[0], div, a[2], b[2]))
    rows.sort(key=lambda r: sd(r[3][3], r[3][0]) - sd(r[2][3], r[2][0]))
    total = tuple([sum(r[i][k] for r in rows) for k in range(5)] for i in (2, 3))
    return rows, total


def day_table(rows, total, detailed):
    idx, heads = ([0, 1, 2, 3, 4], ["Received", "Same Day Invoiced", "Delivered", "Redirected", "Returned"]) \
        if detailed else ([0, 3], ["Received", "Redirected"])
    heads = heads + ["Redirection %"]
    table = []
    lines = [(i + 1, d, n, a, b) for i, (n, d, a, b) in enumerate(rows)] + [("", "Total", TOTAL_LABEL, *total)]
    for sno, div, name, a, b in lines:
        p1, p2 = sd(a[3], a[0]), sd(b[3], b[0])
        table.append([sno, div, name] + [a[j] for j in idx] + [p1] + [b[j] for j in idx] + [p2] + [p2 - p1])
    return heads, table


def write_day_sheet(ws, title, d1, d2, heads, table):
    n = len(heads)
    b1, b2, last = 4, 4 + n, 4 + 2 * n

    def put(r, col, v, fill=None, bold=False, size=11, fmt=None, color=None):
        c = ws.cell(r, col, v)
        c.border, c.font = BOX, Font(name="Calibri", bold=bold, size=size, color=color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if fill:
            c.fill = fill
        if fmt:
            c.number_format = fmt

    put(1, 1, title, bold=True, size=18)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    for col, d, fill in ((b1, d1, BLUE), (b2, d2, CREAM)):
        put(1, col, f"{d:%d.%m.%Y}", fill, True, 16)
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + n - 1)
    put(1, last, "", GOOD)
    for j, h in enumerate(["S.No.", "Division Name", "Office Name"]):
        put(2, 1 + j, h, YELLOW, True)
    for b, fill in ((b1, BLUE), (b2, CREAM)):
        for j, h in enumerate(heads):
            put(2, b + j, h, fill, True)
    put(2, last, "% Inc/Dec", BAD, True, color="C00000")
    for k, row in enumerate(table):
        r, is_total = 3 + k, k == len(table) - 1
        for j, v in enumerate(row):
            col = j + 1
            fill = YELLOW if is_total else (BLUE if b1 <= col < b2 else CREAM if b2 <= col < last else None)
            fmt = "0.00%" if col in (b1 + n - 1, b2 + n - 1, last) else ("0" if col >= b1 else None)
            color = None
            if col == last:
                fill, color = (GOOD, "006100") if v < 0 else (BAD, "9C0006") if v > 0 else (fill, None)
            put(r, col, v, fill, is_total, fmt=fmt, color=color)
    ws.row_dimensions[1].height = ws.row_dimensions[2].height = 38
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width, ws.column_dimensions["C"].width = 7, 24, 26
    for col in range(4, last + 1):
        ws.column_dimensions[L(col)].width = 13
    ws.freeze_panes = "D3"


def day_workbook(title, d1, d2, rows, total):
    wb = Workbook()
    wb.remove(wb.active)
    for name, detailed in (("Summary", False), ("Detailed", True)):
        heads, table = day_table(rows, total, detailed)
        write_day_sheet(wb.create_sheet(name), title, d1, d2, heads, table)
    return to_bytes(wb)


# ---------- 1c-2) The 26 offices (H.O. / S.O. list) and the Missent comparison layout ----------
OFFICES_26 = ["UPPAL H.O", "Sangareddy H.O", "C.B.I.T S.O", "Chitrapuri Colony S.O.", "Golconda S.O",
              "Hydershahkote S.O", "Amberpet S.O", "Hayatnagar S.O", "Hyderguda S.O", "Kanchanbagh S.O",
              "Kattedan I.E. S.O", "P and T Colony (Gaddiannaram) S.O", "Ragannaguda X Road S.O",
              "Rajendranagar S.O", "Rama Krishnapuram S.O", "S.V.P.N.P.A. S.O", "Saroornagar S.O",
              "Vaishalinagar S.O", "Bolarum S.O", "Chandanagar S.O", "Ecil S.O", "Ghatkesar S.O",
              "Malkajgiri S.O", "Miyapur S.O", "Ramakrishna Puram S.O", "Sainikpuri S.O"]
LIST_GROUP = "26 Offices (list below)"
FORMATS = ["Redirection - % on Received (Summary + Detailed sheets)",
           "Missent - % on Invoiced (one sheet, as in the 26 offices sample)"]
MISSENT_SORT = ["List order (as in the sample)", "Increase/Decrease (biggest improvement first)",
                "Office name (A-Z)"]
PINK, KHAKI = PatternFill("solid", fgColor="F8D7DA"), PatternFill("solid", fgColor="EAF1DD")


def okey(n):
    """Office name -> matching key: ignores case, dots and extra spaces ('S.O.' = 'S.O'), keeps the word breaks."""
    return " ".join(re.sub(r"[^a-z0-9 ]", "", str(n).lower()).split())


def pick_list(offices, names):
    """Offices of a given list, kept in list order -> ({key: (name, division, counts)}, [names not found])."""
    have = {}
    for n, dv, c in offices:
        have.setdefault(okey(n), (n, dv, c))
    out, missing = {}, []
    for nm in names:
        k = okey(nm)
        if k in out:
            continue
        if k in have:
            out[k] = have[k]
        else:
            out[k] = (nm, "", [0.0] * 5)
            missing.append(nm)
    return out, missing


def missent_rows(g1, g2, sort):
    """-> (rows, total); row = (office, [invoiced1, redirected1, %missent1, invoiced2, redirected2, %missent2, change])."""
    zero = (None, "", [0.0] * 5)
    rows = []
    for k in list(g1) + [k for k in g2 if k not in g1]:
        a, b = g1.get(k, zero), g2.get(k, zero)
        i1, r1, i2, r2 = a[2][1], a[2][3], b[2][1], b[2][3]
        p1, p2 = sd(r1, i1), sd(r2, i2)
        rows.append((b[0] or a[0], [i1, r1, p1, i2, r2, p2, p2 - p1]))
    if sort == MISSENT_SORT[1]:
        rows.sort(key=lambda r: r[1][6])
    elif sort == MISSENT_SORT[2]:
        rows.sort(key=lambda r: r[0].lower())
    i1, r1 = sum(r[1][0] for r in rows), sum(r[1][1] for r in rows)
    i2, r2 = sum(r[1][3] for r in rows), sum(r[1][4] for r in rows)
    p1, p2 = sd(r1, i1), sd(r2, i2)
    return rows, ("Total", [i1, r1, p1, i2, r2, p2, p2 - p1])


def write_missent_sheet(ws, title, lab1, lab2, rows, total):
    def put(r, col, v, fill=None, bold=False, fmt=None, color=None, size=11, h="center"):
        c = ws.cell(r, col, v)
        c.border, c.font = BOX, Font(name="Times New Roman", bold=bold, size=size, color=color)
        c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=True)
        if fill:
            c.fill = fill
        if fmt:
            c.number_format = fmt

    for col in range(1, 10):
        put(1, col, title if col == 1 else None, None, True, size=14)
    ws.merge_cells("A1:I1")
    for col, h in ((1, "S.no"), (2, "Office Name"), (9, "Increase/ Decrease")):
        put(2, col, h, None, True)
        put(3, col, None)
        ws.merge_cells(start_row=2, start_column=col, end_row=3, end_column=col)
    for col0, lab, fill in ((3, lab1, PINK), (6, lab2, KHAKI)):
        for k, h in enumerate(("Same Day Invoiced", "D0 Redirected", "%Missent")):
            put(2, col0 + k, lab if k == 0 else None, fill, True, size=12)
            put(3, col0 + k, h, fill, True)
        ws.merge_cells(start_row=2, start_column=col0, end_row=2, end_column=col0 + 2)
    for k, (name, vals) in enumerate(rows + [total]):
        is_total, r = k == len(rows), 4 + k
        put(r, 1, "Total" if is_total else k + 1, None, is_total)
        put(r, 2, "" if is_total else name, None, is_total, h="center")
        if is_total:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        for j, v in enumerate(vals):
            col = 3 + j
            fill = PINK if col <= 5 else KHAKI if col <= 8 else None
            color = None
            if col == 9:
                color = "FF0000" if v > 1e-9 else "00B050" if v < -1e-9 else None
            put(r, col, v, fill, is_total, "0.00%" if col in (5, 8, 9) else "0", color)
    ws.row_dimensions[1].height, ws.row_dimensions[2].height, ws.row_dimensions[3].height = 26, 32, 36
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 7, 38
    for col in "CDEFGHI":
        ws.column_dimensions[col].width = 14
    ws.freeze_panes = "C4"


def missent_workbook(title, lab1, lab2, rows, total):
    wb = Workbook()
    ws = wb.active
    ws.title = "Missent Comparison"
    write_missent_sheet(ws, title, lab1, lab2, rows, total)
    return to_bytes(wb)


# ---------- 1d) HQ D+0 reports (Division / Sub Division / Office level, from Office Level exports) ----------
LEVELS = {"Division-wise": "Division", "Sub Division-wise": "Sub Division", "Office-wise": "Office"}
LEVEL_PARENTS = {"Division": [], "Sub Division": ["Division"], "Office": ["Division", "Sub Division"]}
LEVEL_SHORT = {"Division": "Div", "Sub Division": "Sub Div", "Office": "Office"}
LEVEL_TITLE = {"Division": "", "Sub Division": "Sub Division wise - ", "Office": "Office wise - "}
HQ_FMT = ["#,##0", "#,##0", "0.00%", "#,##0", "0.00%", "#,##0", "0.00%", "#,##0", "0.00%", "0.00%", "#,##0", "0.00%"]
SORT_OPTIONS = {"Name (Division / Sub Division / Office)": -1, "Received": 0, "Invoiced": 1, "Invoiced %": 2,
                "D+0 Delivered": 3, "D+0 Delivered %": 4, "D+0 Redirected": 5, "D+0 Redirected %": 6,
                "D+0 Returned": 7, "D+0 Returned %": 8, "D+0 (Delivered+Returned+Redirected) %": 9,
                "Return + Redirection": 10, "Return + Redirection %": 11}
SORT_DEFAULT = "Return + Redirection %"
ORDERS = ["Ascending (low to high)", "Descending (high to low)"]


def clean(x):
    return " ".join(str(x).split()) if pd.notna(x) else ""


def read_hq(f):
    """Office Level export -> list of {id, div, sub, name, c=[5 counts]}."""
    data = io.BytesIO(f.getvalue())
    df = pd.read_csv(data) if f.name.lower().endswith(".csv") else pd.read_excel(data)
    miss = [c for c in ["Division Name", "Office Name"] + RAW_COLS[1:] if c not in df.columns]
    if miss:
        raise ValueError(f"{f.name}: these columns are missing: {', '.join(miss)}. "
                         "Please upload the raw export from the portal.")
    if "Sub Division Name" not in df.columns:
        df["Sub Division Name"] = ""
        rank = 0                                   # Division Level export
    elif (df["Office Name"] == df["Sub Division Name"]).all():
        rank = 1                                   # Sub Division Level export
    elif (df["Office Name"] == df["Division Name"]).all():
        rank = 0
    else:
        rank = 2                                   # Office Level export
    for c in RAW_COLS[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    rows = []
    for i, r in enumerate(df.to_dict("records")):
        div, sub, name = clean(r["Division Name"]), clean(r["Sub Division Name"]), clean(r["Office Name"])
        oid = str(r["office_id"]) if "office_id" in r and pd.notna(r["office_id"]) else f"{div}|{sub}|{name}|{i}"
        rows.append({"id": oid, "div": div, "sub": sub, "name": name, "c": [float(r[c]) for c in RAW_COLS[1:]]})
    return rows, rank


def hq_aggregate(rows, level):
    """-> {key: [display name, [parent names], counts]}"""
    out = {}
    for r in rows:
        if level == "Division":
            key, name, par = r["div"], r["div"], []
        elif level == "Sub Division":
            sub = r["sub"] or f"{r['div']} (no sub division)"
            key, name, par = (r["div"], sub), sub, [r["div"]]
        else:
            key, name, par = r["id"], r["name"], [r["div"], r["sub"]]
        if key in out:
            out[key][2] = [a + b for a, b in zip(out[key][2], r["c"])]
        else:
            out[key] = [name, par, list(r["c"])]
    return out


def hq_add(ds):
    out = {}
    for d in ds:
        for k, v in d.items():
            if k in out:
                out[k][2] = [a + b for a, b in zip(out[k][2], v[2])]
            else:
                out[k] = [v[0], v[1], list(v[2])]
    return out


def hq_sub(a, b):
    out = {k: [v[0], v[1], [x - y for x, y in zip(v[2], b[k][2] if k in b else [0.0] * 5)]] for k, v in a.items()}
    for k, v in b.items():
        if k not in a:
            out[k] = [v[0], v[1], [-x for x in v[2]]]
    return out


def hq_block(title, level, d, recv, sort=11, desc=False):
    """One table. recv=True -> all % on Received, False -> all % on Invoiced. Sorted by Return+Redirection %."""
    on = "On received" if recv else "On invoiced"
    den = (lambda c: c[0]) if recv else (lambda c: c[1])
    items = [v for v in d.values() if any(v[2])]
    items.sort(key=(lambda v: v[0].lower()) if sort < 0 else (lambda v: hq_metrics(v[2], recv)[sort]), reverse=desc)
    npar = len(LEVEL_PARENTS[level])

    def line(sno, name, par, c):
        dn, rr = den(c), c[3] + c[4]
        return ([sno, name] + par + [c[0], c[1], sd(c[1], c[0]), c[2], sd(c[2], dn), c[3], sd(c[3], dn),
                c[4], sd(c[4], dn), sd(c[2] + rr, dn), rr, sd(rr, dn)])

    table = [line(i + 1, v[0], v[1], v[2]) for i, v in enumerate(items)]
    table.append(line("", "Total", [""] * npar, [sum(v[2][k] for v in items) for k in range(5)]))
    heads = (["S.No.", f"Name of the {level}"] + LEVEL_PARENTS[level] +
             ["Received", "Invoiced", "Invoiced %", "D+0 Delivered", f"D+0 Delivered %\n({on})",
              "D+0 Redirected", f"D+0 Redirected %\n({on})", "D+0 Returned", f"D+0 Returned %\n({on})",
              f"D+0 (Delivered+Returned+Redirected) \n({on})", "Return + Redirection", "Return + Redirection\n%"])
    return title, heads, table, [None] * (2 + npar) + HQ_FMT


SHORT_PROD = {"All Products": "AllProd", "Speed Letter": "SpLetter"}


def hq_sheet_name(prefix, product, level, recv, multi_level, both):
    lv = f" - {LEVEL_SHORT[level]}" if multi_level else ""
    bs = (" (Rcvd)" if recv else " (Inv)") if both else ""
    name = prefix + product + lv + bs
    if len(name) > 31:                                             # Excel allows 31 characters only
        lv = f" {LEVEL_SHORT[level].replace(' ', '')}" if multi_level else ""
        bs = (" (R)" if recv else " (I)") if both else ""
        name = prefix + product + lv + bs
    if len(name) > 31:
        name = prefix + SHORT_PROD.get(product, product) + lv + bs
    return name[:31]


def add_hq_report_sheets(wb, jobs, levels, bases, period, sort=11, desc=False, prefix=""):
    """Writes HQ report sheets into wb. jobs = [(sheet, report name, [All Offices row-lists], [BO row-lists],
    file level rank 0/1/2)]; bases = [True (received), False (invoiced)]. Returns (previews, notes)."""
    previews, notes = {}, []
    need = {"Division": 0, "Sub Division": 1, "Office": 2}
    label = {"Division": "any export", "Sub Division": "a Sub Division Level or Office Level export",
             "Office": "an Office Level export"}
    for sheet, rname, all_f, bo_f, rank in jobs:
        for level in levels:
            if need[level] > rank:
                notes.append(f"{rname}: {level}-wise report skipped - it needs {label[level]}.")
                continue
            a = hq_add([hq_aggregate(r, level) for r in all_f])
            scopes = [("All Offices", a)]
            if bo_f:
                b = hq_add([hq_aggregate(r, level) for r in bo_f])
                scopes += [("BOs", b), ("Excluding BOs", hq_sub(a, b))]
            for recv in bases:
                blocks = [hq_block(f"Delivery Transit Analysis {rname}- {LEVEL_TITLE[level]}{sc} -{period}",
                                   level, d, recv, sort, desc) for sc, d in scopes]
                name = hq_sheet_name(prefix, sheet, level, recv, len(levels) > 1, len(bases) > 1)
                ws, r0, npar = wb.create_sheet(name), 1, len(LEVEL_PARENTS[level])
                for title, heads, table, fmts in blocks:
                    write_block(ws, 1, title, heads, table, fmts, r0)
                    r0 += len(table) + 3
                ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 7, 34
                for col in range(3, 3 + npar):
                    ws.column_dimensions[L(col)].width = 30
                for col in range(3 + npar, 15 + npar):
                    ws.column_dimensions[L(col)].width = 14
                ws.freeze_panes = "C1"
                previews[ws.title] = [(t, pd.DataFrame(tb, columns=[h.replace("\n", " ") for h in hd]), fm)
                                      for t, hd, tb, fm in blocks]
    return previews, notes


def make_hq_reports(jobs, levels, bases, d1, d2, sort=11, desc=False):
    wb = Workbook()
    wb.remove(wb.active)
    previews, notes = add_hq_report_sheets(wb, jobs, levels, bases, f"{d1:%d.%m.%Y} to {d2:%d.%m.%Y}", sort, desc)
    if not previews:
        raise ValueError(" ".join(notes) or "Nothing to make.")
    return to_bytes(wb), previews, notes


def period_text(p1, p2):
    if (p2[0] - p1[1]).days == 1:                                  # back-to-back periods -> one range
        return f"{p1[0]:%d.%m.%Y} to {p2[1]:%d.%m.%Y}"
    return f"{p1[0]:%d.%m.%Y} to {p1[1]:%d.%m.%Y} + {p2[0]:%d.%m.%Y} to {p2[1]:%d.%m.%Y}"


# ---------- 1e) HQ D+0 comparison: every column as 1st Period | 2nd Period | Comparison ----------
DIRN = [0, 0, 1, 0, 1, 0, -1, 0, -1, 1, 0, -1]     # +1 higher is better, -1 lower is better, 0 neutral
CHG = {"#,##0": "+#,##0;-#,##0;0", "0.00%": "+0.00%;-0.00%;0.00%"}
TRIO_FMT = HQ_FMT + HQ_FMT + [CHG[f] for f in HQ_FMT]      # block 1, block 2, comparison block
FMT_PY = {"#,##0": "{:,.0f}", "0.00%": "{:.2%}", CHG["#,##0"]: "{:+,.0f}", CHG["0.00%"]: "{:+.2%}"}
GREY = PatternFill("solid", fgColor="E7E6E6")
HQ_GROUPS = [("All Products", "All Products", ["ap_all"], ["ap_bo"], 1),
             ("Speed Letter", "Speed Post Document", ["sl_all"], ["sl_bo"], 1),
             ("Parcels", "Consolidated Parcels", ["sp_all", "rp_all"], ["sp_bo", "rp_bo"], 2)]


def hq_metrics(c, recv):
    dn, rr = (c[0] if recv else c[1]), c[3] + c[4]
    return [c[0], c[1], sd(c[1], c[0]), c[2], sd(c[2], dn), c[3], sd(c[3], dn), c[4], sd(c[4], dn),
            sd(c[2] + rr, dn), rr, sd(rr, dn)]


def trio_heads(on):
    return ["Received", "Invoiced", "Invoiced %", "D+0 Delivered", f"D+0 Delivered %\n({on})",
            "D+0 Redirected", f"D+0 Redirected %\n({on})", "D+0 Returned", f"D+0 Returned %\n({on})",
            f"D+0 (Delivered+Returned+Redirected)\n({on})", "Return + Redirection",
            f"Return + Redirection %\n({on})"]


def hq_scopes(all_f, bo_f, level):
    a = hq_add([hq_aggregate(r, level) for r in all_f])
    out = [("All Offices", a)]
    if bo_f:
        b = hq_add([hq_aggregate(r, level) for r in bo_f])
        out += [("BOs", b), ("Excluding BOs", hq_sub(a, b))]
    return out


def trio_table(level, d1, d2, recv, sort=11, desc=False, on=1):
    """Rows sorted by the chosen column (on: 0 = 1st period, 1 = 2nd period, 2 = comparison); last row = Total."""
    items = []
    for k in list(d2) + [k for k in d1 if k not in d2]:
        src = d2[k] if k in d2 else d1[k]
        c1, c2 = (d1[k][2] if k in d1 else [0.0] * 5), (d2[k][2] if k in d2 else [0.0] * 5)
        if any(c1) or any(c2):
            items.append((src[0], src[1], c1, c2))
    def key(t):
        if sort < 0:
            return t[0].lower()
        a, b = hq_metrics(t[2], recv)[sort], hq_metrics(t[3], recv)[sort]
        return (a, b, b - a)[on]

    items.sort(key=key, reverse=desc)

    def line(sno, name, par, c1, c2):
        m1, m2 = hq_metrics(c1, recv), hq_metrics(c2, recv)
        return [sno, name] + par + m1 + m2 + [round(y - x, 4) for x, y in zip(m1, m2)]   # 4 decimals = 2 decimals in %

    npar = len(LEVEL_PARENTS[level])
    table = [line(i + 1, *it) for i, it in enumerate(items)]
    table.append(line("", "Total", [""] * npar, [sum(t[2][k] for t in items) for k in range(5)],
                      [sum(t[3][k] for t in items) for k in range(5)]))
    return ["S.No.", f"Name of the {level}"] + LEVEL_PARENTS[level], table


def cell_put(ws, r, col, v, fill=None, bold=False, fmt=None, color=None):
    c = ws.cell(r, col, v)
    c.border, c.font = BOX, Font(name="Calibri", bold=bold, color=color)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


SAGE = PatternFill("solid", fgColor="E9EEDC")
P1_TINT, P2_TINT = PatternFill("solid", fgColor="FFFFCC"), PatternFill("solid", fgColor="F4F7EC")


def write_trio(ws, r0, title, name_heads, groups, table, tops):
    """Row r0 = title, r0+1 = Period 1 / Period 2 / comparison labels (with dates), r0+2 = metric names.
    Columns: names | Period 1 metrics | Period 2 metrics | comparison of every metric."""
    npar, nm = len(name_heads), len(groups)
    ncols = npar + 3 * nm
    for col in range(1, ncols + 1):
        cell_put(ws, r0, col, title if col == 1 else None, GREEN, True)
    ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=ncols)
    for j, h in enumerate(name_heads):
        cell_put(ws, r0 + 1, j + 1, h, None, True)
        cell_put(ws, r0 + 2, j + 1, None)
        ws.merge_cells(start_row=r0 + 1, start_column=j + 1, end_row=r0 + 2, end_column=j + 1)
    fills = (YELLOW, SAGE, None)
    for b in range(3):
        c0 = npar + 1 + b * nm
        for k, h in enumerate(groups):
            cell_put(ws, r0 + 1, c0 + k, tops[b] if k == 0 else None, fills[b], True)
            cell_put(ws, r0 + 2, c0 + k, h if b < 2 else h.replace("\n", " ") + " Comparison", fills[b], True)
        ws.merge_cells(start_row=r0 + 1, start_column=c0, end_row=r0 + 1, end_column=c0 + nm - 1)
    for k, row in enumerate(table):
        last = k == len(table) - 1
        for j, v in enumerate(row):
            fill = YELLOW if last else None
            fmt = color = None
            if j >= npar:
                b, g = divmod(j - npar, nm)
                fmt = TRIO_FMT[j - npar]
                if not last:
                    fill = (P1_TINT, P2_TINT, None)[b]
                if b == 2 and DIRN[g] and abs(v) > 1e-9:
                    fill, color = (GOOD, "006100") if v * DIRN[g] > 0 else (BAD, "9C0006")
            c = cell_put(ws, r0 + 3 + k, j + 1, v, fill, last, fmt, color)
            if 0 < j < npar:
                c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r0 + 1].height = 24
    ws.row_dimensions[r0 + 2].height = 100


def add_hq_comparison_sheets(wb, jobs, levels, bases, labels, ranges, sort=11, desc=False, on=1, prefix=""):
    """jobs = [(sheet, report name, all1, bo1, all2, bo2, rank)]; labels = names of the two compared periods;
    ranges = date text of the two periods (shown at the top of each block). Returns (previews, notes)."""
    previews, notes = {}, []
    need = {"Division": 0, "Sub Division": 1, "Office": 2}
    label = {"Division": "any export", "Sub Division": "a Sub Division Level or Office Level export",
             "Office": "an Office Level export"}
    for sheet, rname, a1, b1, a2, b2, rank in jobs:
        if bool(b1) != bool(b2):
            notes.append(f"{rname}: BOs file given for one of the compared periods only - "
                         "BOs / Excluding BOs tables skipped.")
        for level in levels:
            if need[level] > rank:
                notes.append(f"{rname}: {level}-wise report skipped - it needs {label[level]}.")
                continue
            s1, s2 = dict(hq_scopes(a1, b1, level)), dict(hq_scopes(a2, b2, level))
            scopes = [n for n in s1 if n in s2]
            for recv in bases:
                groups = trio_heads("On received" if recv else "On invoiced")
                name = hq_sheet_name(prefix, sheet, level, recv, len(levels) > 1, len(bases) > 1)
                ws, r0, blocks = wb.create_sheet(name), 1, []
                for sc in scopes:
                    title = (f"Delivery Transit Analysis {rname}- {LEVEL_TITLE[level]}{sc} - "
                             f"{labels[0]} {ranges[0]} vs {labels[1]} {ranges[1]}")
                    tops = (f"{labels[0]} : {ranges[0]}", f"{labels[1]} : {ranges[1]}", f"{labels[0]} Vs {labels[1]}")
                    nh, table = trio_table(level, s1[sc], s2[sc], recv, sort, desc, on)
                    write_trio(ws, r0, title, nh, groups, table, tops)
                    r0 += len(table) + 4
                    flat = [g.replace(chr(10), " ") for g in groups]
                    cols = (nh + [f"{labels[0]} - {g}" for g in flat] + [f"{labels[1]} - {g}" for g in flat]
                            + [f"{g} Comparison" for g in flat])
                    blocks.append((title, pd.DataFrame(table, columns=cols), [None] * len(nh) + TRIO_FMT))
                ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 7, 34
                npar = len(LEVEL_PARENTS[level])
                for col in range(3, 3 + npar):
                    ws.column_dimensions[L(col)].width = 30
                for col in range(3 + npar, 3 + npar + 36):
                    ws.column_dimensions[L(col)].width = 14
                ws.freeze_panes = "C1"
                previews[ws.title] = blocks
    return previews, notes


def make_hq_workbook(P, D, levels, bases, sort=11, desc=False, on=1):
    """P = {1: .., 2: .., 3: ..} period data from hq_read_period ({} if none); D = {k: (from, to)}.
    Sheet groups, in this order: each period's report, Period 2 vs Period 3 comparison, Period 2 + 3 addition
    (no comparison), then Period 1 compared with the added periods (or with Period 2 / Period 3 alone)."""
    wb = Workbook()
    wb.remove(wb.active)
    previews, notes = {}, []
    names = {g[0]: g[1] for g in HQ_GROUPS}
    txt = lambda t: f"{t[0]:%d.%m.%Y} to {t[1]:%d.%m.%Y}"
    have = [k for k in (1, 2, 3) if P.get(k)]

    def keep(res):
        previews.update(res[0])
        notes.extend(n for n in res[1] if n not in notes)

    def in_both(pa, pb, what):
        for sh in names:
            if (sh in pa) != (sh in pb):
                notes.append(f"{names[sh]}: files are needed for both ({what}) - skipped in that comparison.")
        return [sh for sh in names if sh in pa and sh in pb]

    def cjobs(pa, pb, what):
        out = []
        for sh in in_both(pa, pb, what):
            x, y = pa[sh], pb[sh]
            out.append((sh, x[0], x[1], x[2], y[1], y[2], min(x[3], y[3])))
        return out

    for k in have:                                                   # 1-3) every period on its own
        jobs = [(sh, v[0], v[1], v[2], v[3]) for sh, v in P[k].items()]
        keep(add_hq_report_sheets(wb, jobs, levels, bases, txt(D[k]), sort, desc, f"P{k} " if len(have) > 1 else ""))
    added = None
    if P.get(2) and P.get(3):
        keep(add_hq_comparison_sheets(wb, cjobs(P[2], P[3], "Period 2 and Period 3"), levels, bases,
                                      ("Period 2", "Period 3"), (txt(D[2]), txt(D[3])),
                                      sort, desc, on, "P2vP3 "))     # 4) Period 2 vs Period 3
        added = {}
        for sh in in_both(P[2], P[3], "Period 2 and Period 3"):      # 5) Period 2 + Period 3
            x, y = P[2][sh], P[3][sh]
            if bool(x[2]) != bool(y[2]):
                notes.append(f"{x[0]}: BOs file given for one period only - no BOs tables in the addition.")
            added[sh] = (x[0], x[1] + y[1], (x[2] + y[2]) if (x[2] and y[2]) else [], min(x[3], y[3]))
        keep(add_hq_report_sheets(wb, [(sh, *v) for sh, v in added.items()], levels, bases,
                                  period_text(D[2], D[3]), sort, desc, "P2+P3 "))
    if P.get(1):                                                     # 6) Period 1 vs (Period 2 + 3, or 2, or 3)
        late = None
        if added is not None:
            late, lab, pre, rng = added, "Period 2+3", "P1v(P2+P3) ", period_text(D[2], D[3])
        elif P.get(2) or P.get(3):
            k = 2 if P.get(2) else 3
            late, lab, pre, rng = P[k], f"Period {k}", f"P1vP{k} ", txt(D[k])
        if late:
            keep(add_hq_comparison_sheets(wb, cjobs(P[1], late, f"Period 1 and {lab}"), levels, bases,
                                          ("Period 1", lab), (txt(D[1]), rng),
                                          sort, desc, on, pre))
    if not previews:
        raise ValueError(" ".join(notes) or "Nothing to make.")
    return to_bytes(wb), previews, notes


# ---------- 1f) WhatsApp texts for NDCs / DCs / 24 (Speed Post) from the daily Excel ----------
TEXT_KINDS = {
    "ndc": dict(head="*NDCs* Delivery Performance dated", who="NDCs", target="Daily Target", overall=True,
                close="Divisions are requested to improve delivery performance in low-performing Offices "
                      "and take action to curtail Return and Redirection"),
    "dc": dict(head="*DCs* Delivery Performance dated", who="DCs", target="Daily Target", overall=True,
               close="Divisions are requested to improve delivery performance in low-performing Offices "
                     "and take action to curtail Return and Redirection"),
    "24": dict(head="*_24 Speed Post_ at DCs/NDCs* Delivery Performance dated", who="NDCs/DCs",
               target="Successful Daily Target", overall=False,
               close="Divisions are requested to improve delivery performance in low-performing Offices "
                     "and take action to curtail Return and Redirection"),
}


def pct2(v):
    """Fraction -> percent rounded to 2 decimals, half up (shows like Excel does: 0.85425 -> 85.43)."""
    return Decimal(repr(round(float(v) * 100, 8))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def num_pct(v):
    return pct2(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def find_sheet(wb, name):
    for n in wb.sheetnames:
        if " ".join(n.lower().split()) == name:
            return wb[n]
    raise ValueError(f"Sheet '{name}' not found in the Excel file.")


def parse_sections(rows):
    """Finds every '... Delivery Performance report dated ...' table on a sheet.
    -> [{kind, title, items: [(office, {role: Decimal %})], total: {role: Decimal %}}]"""
    secs, i = [], 0
    while i < len(rows) - 1:
        title = rows[i][1] if len(rows[i]) > 1 else None
        if not (isinstance(title, str) and "delivery performance" in title.lower()):
            i += 1
            continue
        roles = {}
        for j, h in enumerate(rows[i + 1]):
            t = norm(h) if h is not None else ""
            if t.startswith("%") and "successfuldelivery" in t:
                roles.setdefault("succ", j)
            elif t == "%overalldelivery":
                roles.setdefault("overall", j)
            elif t.startswith("missent") and t.endswith("%"):
                roles.setdefault("redir", j)          # 'Redirections' in the text = Missent %
            elif t.startswith("return") and t.endswith("%"):
                roles.setdefault("ret", j)
        items, total, k = [], {}, i + 2
        while k < len(rows):
            r = rows[k]
            name = r[1] if len(r) > 1 else None
            if not isinstance(name, str) or not name.strip():
                break
            vals = {role: num_pct(r[j]) for role, j in roles.items() if j < len(r)}
            if "overal" in name.lower():              # 'HQR Overal Performance'
                total = vals
                k += 1
                break
            items.append((" ".join(name.split()), vals))
            k += 1
        t = title.strip().lower()
        kind = "ndc" if t.startswith("ndc") else "dc" if t.startswith("dc") else "24" if t.startswith("24") else "other"
        secs.append({"kind": kind, "title": title, "items": items, "total": total})
        i = k
    return secs


def read_text_workbook(data):
    """-> {'ndc': section, 'dc': section, '24': section} (missing ones are left out)."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        out = {}
        for sheet in ("morning report", "24 morning report"):
            try:
                ws = find_sheet(wb, sheet)
            except ValueError:
                continue
            for sec in parse_sections([list(r) for r in ws.iter_rows(values_only=True)]):
                out.setdefault(sec["kind"], sec)
    return out


def report_date(title, fallback):
    m = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", title or "")
    try:
        d = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else fallback
    except ValueError:
        d = fallback
    return f"{d:%d.%m.%Y} ({d:%A})"


def make_text(kind, sec, fallback_date, target, mark, redir, ret):
    """WhatsApp text for one section. mark = 'successful delivery below', redir / ret = 'above' limits (in %)."""
    cfg, g = TEXT_KINDS[kind], (lambda x: f"{x:g}")
    for role in ("succ", "redir", "ret"):
        if not any(v.get(role) is not None for _, v in sec["items"]):
            raise ValueError(f"column for '{role}' not found in the {cfg['who']} table")

    def lines(role, test):
        hit = [(n, v[role]) for n, v in sec["items"] if v.get(role) is not None and test(v[role])]
        hit.sort(key=lambda t: t[1], reverse=True)
        return "\n".join(f"* {n}\t{p}%" for n, p in hit) or "* None"

    show = lambda role: f"{sec['total'][role]}%" if sec["total"].get(role) is not None else "N/A"
    out = [f"{cfg['head']} *{report_date(sec['title'], fallback_date)}*", "",
           f"{cfg['target']} : {g(target)}%", f"* Successful Delivery: *{show('succ')}*"]
    if cfg["overall"]:
        out.append(f"* Overall Delivery: *{show('overall')}*")
    out += ["", f"*{cfg['who']} with Successful Delivery below {g(mark)}% are*",
            lines("succ", lambda p: p < Decimal(str(mark))), "",
            f"*{cfg['who']} with Redirections above {g(redir)}%*", lines("redir", lambda p: p > Decimal(str(redir))), "",
            f"*{cfg['who']} with Returns above {g(ret)}%*", lines("ret", lambda p: p > Decimal(str(ret))), "",
            cfg["close"]]
    return "\n".join(out)


# ---------- 1g) DSS App (Delivery Mobile App Usage) report ----------
DSS_SORT = ["Name (A-Z)", "Period 1 Mobile Usage", "Period 2 Mobile Usage", "% Difference",
            "Period 1 PDM article count", "Period 2 PDM article count"]
LAVENDER, PEACH = PatternFill("solid", fgColor="DDA0DD"), PatternFill("solid", fgColor="F8CBAD")


def read_dss(f):
    """Mobile App Usage export (office level) -> (list of {id, div, sub, name, c=[PDM, DSS]}, level rank 0/1/2)."""
    data = io.BytesIO(f.getvalue())
    df = pd.read_csv(data) if f.name.lower().endswith(".csv") else pd.read_excel(data)
    miss = [c for c in ["division-name", "office_name", "total_pdm_art_count", "total_dss_art_count"]
            if c not in df.columns]
    if miss:
        raise ValueError(f"{f.name}: these columns are missing: {', '.join(miss)}. "
                         "Please upload the Delivery Mobile App Usage export.")
    if "sub-division-name" not in df.columns:
        df["sub-division-name"] = ""
    df = df[~df["office_name"].astype(str).str.strip().str.lower().str.startswith("summary")].copy()
    for c in ("total_pdm_art_count", "total_dss_art_count"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    rows = []
    for i, r in enumerate(df.to_dict("records")):
        div, sub, name = clean(r["division-name"]), clean(r["sub-division-name"]), clean(r["office_name"])
        oid = str(r["office_id"]) if "office_id" in r and pd.notna(r["office_id"]) else f"{div}|{sub}|{name}|{i}"
        rows.append({"id": oid, "div": div, "sub": sub, "name": name,
                     "c": [float(r["total_pdm_art_count"]), float(r["total_dss_art_count"])]})
    if not rows:
        raise ValueError(f"{f.name}: no data rows found.")
    if all(r["sub"] == "" for r in rows) or all(r["name"] == r["div"] for r in rows):
        rank = 0                                           # Division Level export
    elif all(r["name"] == r["sub"] for r in rows):
        rank = 1                                           # Sub Division Level export
    else:
        rank = 2                                           # Office Level export
    return rows, rank


def dss_aggregate(rows, level):
    out = {}
    for r in rows:
        if level == "Division":
            key, name, par = r["div"], r["div"], []
        elif level == "Sub Division":
            key, name, par = (r["div"], r["sub"]), r["sub"], [r["div"]]
        else:
            key, name, par = r["id"], r["name"], [r["div"], r["sub"]]
        if key in out:
            out[key][2] = [a + b for a, b in zip(out[key][2], r["c"])]
        else:
            out[key] = [name, par, list(r["c"])]
    return out


def dss_table(d1, d2, hide_zero, sort, desc):
    """-> (rows, total). row = (name, parents, c1, c2, usage1, usage2, difference); c = [PDM, DSS]."""
    rows = []
    for k in list(d1) + [k for k in d2 if k not in d1]:
        a, b = d1.get(k), d2.get(k)
        src = b or a
        c1, c2 = (a[2] if a else [0.0, 0.0]), (b[2] if b else [0.0, 0.0])
        if hide_zero and not any(c1) and not any(c2):
            continue
        u1, u2 = sd(c1[1], c1[0]), sd(c2[1], c2[0])
        rows.append((src[0], src[1], c1, c2, u1, u2, u2 - u1))
    keys = {"Name (A-Z)": lambda r: r[0].lower(), "Period 1 Mobile Usage": lambda r: r[4],
            "Period 2 Mobile Usage": lambda r: r[5], "% Difference": lambda r: r[6],
            "Period 1 PDM article count": lambda r: r[2][0], "Period 2 PDM article count": lambda r: r[3][0]}
    rows.sort(key=keys[sort], reverse=desc)
    t1, t2 = [sum(r[2][i] for r in rows) for i in (0, 1)], [sum(r[3][i] for r in rows) for i in (0, 1)]
    u1, u2 = sd(t1[1], t1[0]), sd(t2[1], t2[0])
    return rows, ("Total", [""] * (len(rows[0][1]) if rows else 0), t1, t2, u1, u2, u2 - u1)


def dss_values(r):
    return [r[2][0], r[2][1], r[4], r[3][0], r[3][1], r[5], r[6]]


def write_dss_sheet(ws, level, label1, label2, rows, total):
    npar = len(LEVEL_PARENTS[level])
    n0 = 1 + npar                                          # name + parent columns
    b1, b2, last = n0 + 1, n0 + 4, n0 + 7
    for j, h in enumerate([f"Name of the {level}"] + LEVEL_PARENTS[level]):
        cell_put(ws, 1, j + 1, h, LAVENDER, True)
        cell_put(ws, 2, j + 1, None, LAVENDER)
        ws.merge_cells(start_row=1, start_column=j + 1, end_row=2, end_column=j + 1)
    for col, lab in ((b1, label1), (b2, label2)):
        for k in range(3):
            cell_put(ws, 1, col + k, lab if k == 0 else None, LAVENDER, True)
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        for k, h in enumerate(("Total PDM article count", "Total DSS article count", "Mobile Usage")):
            cell_put(ws, 2, col + k, h, PEACH, True)
    cell_put(ws, 1, last, "% Difference", LAVENDER, True)
    cell_put(ws, 2, last, None, LAVENDER)
    ws.merge_cells(start_row=1, start_column=last, end_row=2, end_column=last)
    for k, r in enumerate(rows + [total]):
        is_total, row = k == len(rows), 3 + k
        fill = YELLOW if is_total else None
        c = cell_put(ws, row, 1, r[0], fill, is_total)
        c.alignment = Alignment(horizontal="center" if is_total else "left", vertical="center")
        for j, p in enumerate(r[1]):
            c = cell_put(ws, row, 2 + j, p, fill, is_total)
            c.alignment = Alignment(horizontal="left", vertical="center")
        for j, v in enumerate(dss_values(r)):
            col = b1 + j
            color = None
            if col == last:
                color = "FF0000" if v < -1e-9 else "00B050" if v > 1e-9 else None
            cell_put(ws, row, col, v, fill, is_total, "0.00%" if col in (b1 + 2, b2 + 2, last) else "0", color)
    ws.row_dimensions[1].height, ws.row_dimensions[2].height = 32, 50
    ws.column_dimensions["A"].width = 36
    for col in range(2, 2 + npar):
        ws.column_dimensions[L(col)].width = 30
    for col in range(b1, last + 1):
        ws.column_dimensions[L(col)].width = 17
    ws.freeze_panes = ws.cell(3, 2)


def make_dss_report(rows1, rank1, rows2, rank2, levels, lab1, lab2, hide_zero, sort, desc):
    wb = Workbook()
    wb.remove(wb.active)
    previews, notes = {}, []
    need = {"Division": 0, "Sub Division": 1, "Office": 2}
    for level in levels:
        if need[level] > min(rank1, rank2):
            notes.append(f"{level}-wise report skipped - it needs a "
                         f"{'Sub Division Level or Office Level' if level == 'Sub Division' else 'Office Level'} export.")
            continue
        r, total = dss_table(dss_aggregate(rows1, level), dss_aggregate(rows2, level), hide_zero, sort, desc)
        write_dss_sheet(wb.create_sheet(f"DSS App - {level}"), level, lab1, lab2, r, total)
        cols = ([f"Name of the {level}"] + LEVEL_PARENTS[level]
                + [f"{lab} - {h}" for lab in (lab1, lab2) for h in ("Total PDM article count", "Total DSS article count", "Mobile Usage")]
                + ["% Difference"])
        previews[level] = pd.DataFrame([[x[0]] + x[1] + dss_values(x) for x in r + [total]], columns=cols)
    if not previews:
        raise ValueError(" ".join(notes))
    return to_bytes(wb), previews, notes


# ---------- 1h) COD Collection report (cash vs digital COD transactions) ----------
COD_HEADS = ["Sum of COD Cash Count", "Sum of COD Digital Count", "Sum of Total COD Count", "% of Digital Transactions"]
COD_PARENTS = {"Division": [], "Sub Division": [], "Office": ["Division", "Sub Division"]}
COD_SORT = ["Difference", "% of Digital Transactions", "Name (A-Z)", "COD Cash Count", "COD Digital Count",
            "Total COD Count"]


def read_cod(f):
    """COD Collection export -> (list of {id, div, sub, name, c=[cash, digital, total]}, level rank 0/1/2)."""
    data = io.BytesIO(f.getvalue())
    df = pd.read_csv(data) if f.name.lower().endswith(".csv") else pd.read_excel(data)
    counts = ["COD Cash Count", "COD Digital Count", "Total COD Count"]
    miss = [c for c in ["Division Name"] + counts if c not in df.columns]
    if miss:
        raise ValueError(f"{f.name}: these columns are missing: {', '.join(miss)}. "
                         "Please upload the COD Collection Performance export.")
    for c in ("Sub Division Name", "Office Name"):
        if c not in df.columns:
            df[c] = ""
    for c in counts:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    blank = lambda x: (x.isna() | (x.astype(str).str.strip() == "")).all()
    off_blank, sub_blank = blank(df["Office Name"]), blank(df["Sub Division Name"])
    rows = []
    for i, r in enumerate(df.to_dict("records")):
        div, sub, name = clean(r["Division Name"]), clean(r["Sub Division Name"]), clean(r["Office Name"])
        try:
            oid = str(int(float(r["Office ID"])))
        except (KeyError, TypeError, ValueError):
            oid = f"{div}|{sub}|{name}"
        rows.append({"id": oid, "div": div, "sub": sub, "name": name or sub or div,
                     "c": [float(r[c]) for c in counts]})
    if not rows:
        raise ValueError(f"{f.name}: no data rows found.")
    if (off_blank and sub_blank) or (not off_blank and all(r["name"] == r["div"] for r in rows)):
        rank = 0                                           # Division Level export
    elif off_blank or all(r["name"] == r["sub"] for r in rows):
        rank = 1                                           # Sub Division Level export
    else:
        rank = 2                                           # Office Level export
    return rows, rank


def cod_aggregate(rows, level):
    out = {}
    for r in rows:
        if level == "Division":
            key, name, par = r["div"], r["div"], []
        elif level == "Sub Division":
            sub = r["sub"] or f"{r['div']} (no sub division)"
            key, name, par = (r["div"], sub), sub, []
        else:
            key, name, par = r["id"], r["name"], [r["div"], r["sub"]]
        if key in out:
            out[key][2] = [a + b for a, b in zip(out[key][2], r["c"])]
        else:
            out[key] = [name, par, list(r["c"])]
    return out


def cod_rows(dicts, sort, desc):
    """dicts = [period 1] or [period 1, period 2] -> (rows, total).
    row = (name, parents, [[cash, digital, total, % digital] per period], difference or None)."""
    two = len(dicts) == 2
    keys = list(dicts[0]) + ([k for k in dicts[1] if k not in dicts[0]] if two else [])
    pct = lambda c: sd(c[1], c[2]) * 100
    rows = []
    for k in keys:
        src = next(d[k] for d in dicts if k in d)
        blocks = []
        for d in dicts:
            c = d[k][2] if k in d else [0.0, 0.0, 0.0]
            blocks.append(c + [pct(c)])
        diff = blocks[1][3] - blocks[0][3] if two and blocks[0][2] and blocks[1][2] else None
        rows.append((src[0], src[1], blocks, diff))
    key = {"Name (A-Z)": lambda r: r[0].lower(), "COD Cash Count": lambda r: r[2][-1][0],
           "COD Digital Count": lambda r: r[2][-1][1], "Total COD Count": lambda r: r[2][-1][2],
           "% of Digital Transactions": lambda r: r[2][-1][3],
           "Difference": (lambda r: r[3]) if two else (lambda r: r[2][-1][3])}[sort]
    ok = sorted([r for r in rows if key(r) is not None], key=key, reverse=desc)
    rows = ok + [r for r in rows if key(r) is None]       # rows without a difference go last
    tot = []
    for p in range(len(dicts)):
        c = [sum(r[2][p][i] for r in rows) for i in range(3)]
        tot.append(c + [pct(c)])
    diff = tot[1][3] - tot[0][3] if two and tot[0][2] and tot[1][2] else None
    return rows, ("Grand Total", [""] * (len(rows[0][1]) if rows else 0), tot, diff)


def write_cod_sheet(ws, level, labels, rows, total):
    """Yellow header like the sample: date over each block of 4 columns, then the column names."""
    npar, two = len(COD_PARENTS[level]), len(labels) == 2
    n0 = 1 + npar
    for j, h in enumerate([level] + COD_PARENTS[level]):
        cell_put(ws, 1, j + 1, h, YELLOW, True)
        cell_put(ws, 2, j + 1, None, YELLOW)
        ws.merge_cells(start_row=1, start_column=j + 1, end_row=2, end_column=j + 1)
    for p, lab in enumerate(labels):
        c0 = n0 + 1 + 4 * p
        for k, h in enumerate(COD_HEADS):
            cell_put(ws, 1, c0 + k, lab if k == 0 else None, YELLOW, True)
            cell_put(ws, 2, c0 + k, h, YELLOW, True)
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 3)
    if two:
        cell_put(ws, 1, n0 + 9, "Difference", YELLOW, True)
        cell_put(ws, 2, n0 + 9, None, YELLOW)
        ws.merge_cells(start_row=1, start_column=n0 + 9, end_row=2, end_column=n0 + 9)
    for k, r in enumerate(rows + [total]):
        is_total, row = k == len(rows), 3 + k
        fill = YELLOW if is_total else None
        cell_put(ws, row, 1, r[0], fill, is_total)
        for j, p in enumerate(r[1]):
            cell_put(ws, row, 2 + j, p, fill, is_total)
        flat = [v for blk in r[2] for v in blk] + ([r[3]] if two else [])
        for j, v in enumerate(flat):
            cell_put(ws, row, n0 + 1 + j, v, fill, is_total, "0.00" if (j % 4 == 3 or j == 8) else "0")
    ws.row_dimensions[1].height, ws.row_dimensions[2].height = 30, 48
    ws.column_dimensions["A"].width = 38
    for col in range(2, 2 + npar):
        ws.column_dimensions[L(col)].width = 30
    for col in range(n0 + 1, n0 + 10):
        ws.column_dimensions[L(col)].width = 17
    ws.freeze_panes = ws.cell(3, 2)


def make_cod_report(p1, p2, levels, sort, desc):
    """p1 / p2 = (rows, rank, date label) or None. One period -> its report; both -> Period 1, Period 2 and
    their comparison (3 sheets per level)."""
    wb = Workbook()
    wb.remove(wb.active)
    previews, notes = {}, []
    per = [(k, p) for k, p in ((1, p1), (2, p2)) if p]
    need = {"Division": 0, "Sub Division": 1, "Office": 2}
    for level in levels:
        if need[level] > min(p[1] for _, p in per):
            notes.append(f"{level}-wise report skipped - it needs "
                         f"{'a Sub Division Level or Office Level' if level == 'Sub Division' else 'an Office Level'} export.")
            continue
        suffix = f" - {LEVEL_SHORT[level]}" if len(levels) > 1 else ""
        aggs = [(k, p[2], cod_aggregate(p[0], level)) for k, p in per]
        sheets = [(f"Period {k}{suffix}", [lab], [d]) for k, lab, d in aggs]
        if len(aggs) == 2:
            sheets.append((f"Comparison{suffix}", [aggs[0][1], aggs[1][1]], [aggs[0][2], aggs[1][2]]))
        for name, labels, ds in sheets:
            rows, total = cod_rows(ds, sort, desc)
            write_cod_sheet(wb.create_sheet(name), level, labels, rows, total)
            cols = [level] + COD_PARENTS[level]
            cols += [f"{lab} - {h}" if len(labels) == 2 else h for lab in labels for h in COD_HEADS]
            cols += ["Difference"] if len(labels) == 2 else []
            table = [[r[0]] + r[1] + [v for blk in r[2] for v in blk] + ([r[3]] if len(labels) == 2 else [])
                     for r in rows + [total]]
            previews[name] = (f"{level}-wise: {' vs '.join(labels)}", pd.DataFrame(table, columns=cols))
    if not previews:
        raise ValueError(" ".join(notes))
    return to_bytes(wb), previews, notes


# ---------- 1i) PowerPoint slides from the HQ comparison (add-on: does not change the Excel reports) ----------
PPT_PRODUCT = {"All Products": "All Products", "Speed Letter": "Speed Post Document", "Parcels": "Parcels"}
PPT_TAG = {"All Offices": "All Offices", "BOs": "BOs", "Excluding BOs": "Exc. BOs"}
PPT_LEVELS = {"Sub Division-wise": "Sub Division", "Division-wise": "Division", "Office-wise": "Office"}
PPT_NAME_W = {("delivery", 1): [2.6], ("redrt", 1): [2.6], ("delivery", 2): [2.2, 2.0], ("redrt", 2): [1.75, 1.65],
              ("delivery", 3): [3.0, 1.9, 1.7], ("redrt", 3): [2.7, 1.8, 1.5]}
PPT_YELLOW, PPT_PEACH = "FFFF00", "F7DCC2"
PPT_RED, PPT_GREEN = "FF0000", "13803A"


def indian(n):
    """312920 -> '3,12,920' (Indian digit grouping, as in the reference slides)."""
    s = str(int(round(n)))
    neg, s = s.startswith("-"), s.lstrip("-")
    if len(s) > 3:
        head, tail, parts = s[:-3], s[-3:], []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ",".join(parts + [tail])
    return ("-" if neg else "") + s


def pstr(v):
    d = pct2(v)
    return f"{Decimal('0.00') if d == 0 else d}%"


def bench_val(txt):
    """'85' / '85%' / '' -> '85' (None when empty)."""
    t = str(txt or "").replace("%", "").strip()
    if not t:
        return None
    try:
        return f"{float(t):g}"
    except ValueError:
        raise ValueError(f"Benchmark '{txt}' is not a number")


def hq_ppt_pairs(P, D):
    """Comparisons that can be put on slides -> {label: (earlier products, later products, earlier dates, later dates)}.
    P = {1: .., 2: .., 3: ..} from hq_read_period; D = {k: (from, to)}."""
    txt = lambda t: f"{t[0]:%d.%m.%Y} to {t[1]:%d.%m.%Y}"
    added = None
    if P.get(2) and P.get(3):
        added = {}
        for sh in P[2]:
            if sh in P[3]:
                x, y = P[2][sh], P[3][sh]
                added[sh] = (x[0], x[1] + y[1], (x[2] + y[2]) if (x[2] and y[2]) else [], min(x[3], y[3]))
    out = {}
    if P.get(1) and added:
        out["Period 1 vs Period 2+3 (added)"] = (P[1], added, txt(D[1]), period_text(D[2], D[3]))
    for a, b in ((1, 2), (1, 3), (2, 3)):
        if P.get(a) and P.get(b):
            out[f"Period {a} vs Period {b}"] = (P[a], P[b], txt(D[a]), txt(D[b]))
    return out


def ppt_name(n):
    """Names as on the reference slides: 'IP MEDAK NORTH Sub Division' -> 'MEDAK NORTH', 'Medak Division' -> 'Medak'."""
    n = re.sub(r"\s*\(no sub division\)\s*$", "", str(n), flags=re.I)
    n = re.sub(r"\s+(sub\s+)?division\s*$", "", n, flags=re.I)
    return re.sub(r"^(IP|ASP)\s+(?=\S)", "", n.strip())


def _ppt_cell(cell, text, size, bold=False, fill="FFFFFF", color="000000", align=None):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.oxml.ns import qn
    from lxml import etree
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):                 # thin black borders (before the fill element)
        ln = etree.SubElement(tcPr, qn(tag), w="9525", cap="flat", cmpd="sng", algn="ctr")
        etree.SubElement(etree.SubElement(ln, qn("a:solidFill")), qn("a:srgbClr"), val="000000")
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor.from_string(fill)
    cell.margin_left = cell.margin_right = Inches(0.04)
    cell.margin_top = cell.margin_bottom = Inches(0.02)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.text_frame.word_wrap = True
    p = cell.text_frame.paragraphs[0]
    p.alignment = align or PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size, run.font.bold, run.font.name = Pt(size), bold, "Calibri"
    run.font.color.rgb = RGBColor.from_string(color)


def _ppt_slide(prs, title, tag, bench, labels, level, nh, rows, total, kind, truncated):
    """One slide: peach title box + scope tag (+ benchmark box), then the yellow-header table of the reference slides.
    rows = data rows of this slide, total = Total row (last slide of a table only) or None."""
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.oxml.ns import qn
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    delivery, npar = kind == "delivery", len(nh)
    parents = list(nh[2:])[::-1]                                          # child -> parent: Office | Sub Division | Division
    nn = 1 + len(parents)
    tag_w, gap, tot_w = 1.85, 0.12, 12.55
    btxt, bsize, bw = "", 12, 0.0
    if bench:
        btxt, bsize = (f"Benchmark: {bench}%", 14) if delivery else (f"Benchmark (Return+Redirection): {bench}%", 12)
        bw = len(btxt) * bsize * 0.0068 + 0.35
    title_w = tot_w - tag_w - gap - ((bw + gap) if bench else 0)
    tsize = 20
    while tsize > 14 and len(title) * tsize * 0.0063 > title_w - 0.3:
        tsize -= 1
    x, boxes = 0.39, []
    for text, w, size, align in ((title, title_w, tsize, PP_ALIGN.LEFT), (tag, tag_w, 20, PP_ALIGN.CENTER),
                                 (btxt, bw, bsize, PP_ALIGN.CENTER)):
        if text:
            boxes.append((text, x, w, size, align))
            x += w + gap
    for text, bx, w, size, align in boxes:
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(bx), Inches(0.28), Inches(w), Inches(0.6))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor.from_string(PPT_PEACH)
        box.line.fill.background()
        box.shadow.inherit = False
        style = box._element.find(qn("p:style"))                       # drop the theme style (no shadow / outline)
        if style is not None:
            box._element.remove(style)
        tf = box.text_frame
        tf.vertical_anchor, tf.word_wrap = MSO_ANCHOR.MIDDLE, True
        tf.margin_left = tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size, run.font.bold, run.font.name = Pt(size), True, "Calibri"
        run.font.color.rgb = RGBColor.from_string("000000")
    cnt = level != "Office"                                              # count column left out on Office slides (space)
    if delivery:
        block = ["Invoiced", "D+0 Actual Delivery", "D+0 Actual Delivery %"]
    else:
        block = (["Invoiced", "Redirection %", "Return %"] + (["Redirection + Return count"] if cnt else [])
                 + ["Redirection + Return %"])
    ng, n = len(block), len(rows)
    ncols = nn + 2 * ng + 1
    nrows = n + (1 if total else 0)
    shape = slide.shapes.add_table(nrows + 2, ncols, Inches(0.39), Inches(1.1), Inches(tot_w), Inches(5.5))
    tbl = shape.table
    tbl.first_row = tbl.horz_banding = False
    name_w = PPT_NAME_W[(kind, nn)]
    num_w = (tot_w - sum(name_w)) / (ncols - nn)
    for j in range(ncols):
        tbl.columns[j].width = Inches(name_w[j] if j < nn else num_w)
    size = 16 if n <= 6 else 14 if n <= 10 else 12 if n <= 14 else 11
    if level == "Office":
        size = 11 if n <= 10 else 10
    hsize = size if (delivery or size > 12) else size - 1.5
    if level == "Office":
        hsize = min(hsize, 9 if delivery else 8)
    h0, h1 = 0.34, 0.6 if delivery else 0.85
    per = min(0.5, (7.5 - 1.1 - 0.3 - h0 - h1) / max(nrows, 1))
    tbl.rows[0].height, tbl.rows[1].height = Inches(h0), Inches(h1)
    for r in range(2, nrows + 2):
        tbl.rows[r].height = Inches(per)
    first = ["Office" if level == "Office" else "Sub Division" if level == "Sub Division" else "Division"] + parents
    heads = first + block + block + ["Difference"]
    dates = [""] * nn + [labels[0]] + [""] * (ng - 1) + [labels[1]] + [""] * (ng - 1) + [""]
    for j in range(ncols):
        _ppt_cell(tbl.cell(0, j), dates[j], size + 1, True, PPT_YELLOW)
        _ppt_cell(tbl.cell(1, j), heads[j], hsize, True, PPT_YELLOW)
    for c0 in (nn, nn + ng):
        tbl.cell(0, c0).merge(tbl.cell(0, c0 + ng - 1))
    good = -1 if kind == "redrt" else 1                                  # which direction of change is good
    for i, r in enumerate(rows + ([total] if total else [])):
        last = bool(total) and i == len(rows)
        m1, m2, ch = r[npar:npar + 12], r[npar + 12:npar + 24], r[npar + 24:npar + 36]
        if delivery:
            cells = lambda m: [indian(m[1]), indian(m[3]), pstr(m[4])]
            diff = ch[4]
        else:
            cells = lambda m: ([indian(m[1]), pstr(m[6]), pstr(m[8])] + ([indian(m[10])] if cnt else []) + [pstr(m[11])])
            diff = ch[11]
        vals = cells(m1) + cells(m2)
        if last:
            lead = ["Total (all rows)" if truncated else ""] + [""] * (nn - 1)
        else:
            lead = [ppt_name(x) for x in [r[1]] + list(r[2:npar])[::-1]]
        fill = PPT_YELLOW if last else "FFFFFF"
        for j, v in enumerate(lead + vals):
            _ppt_cell(tbl.cell(2 + i, j), v, size, last, fill, align=PP_ALIGN.LEFT if (j < nn and not last) else None)
        color = "000000" if abs(diff) < 0.00005 else (PPT_GREEN if diff * good > 0 else PPT_RED)
        _ppt_cell(tbl.cell(2 + i, ncols - 1), pstr(diff), size, last, fill, color)
    return slide


def make_hq_ppt(early, later, ranges, level, wording, bench=(None, None), rows_per=18, max_rows=0,
                worst_first=False, max_slides=700):
    """early / later = {sheet: (report name, [All rows], [BO rows], rank)}; ranges = date text of the two periods;
    bench = (delivery benchmark, return+redirection benchmark) as text or None.
    -> (pptx bytes, notes, number of slides). All Delivery slides first, then all Return & Redirection slides; for each
    product: All Offices, BOs, Exc. BOs (the last two only when both periods have a BOs file). Long tables continue on
    the next slide (rows_per rows per slide); max_rows > 0 keeps only the first rows of each table."""
    from pptx import Presentation
    from pptx.util import Inches
    need = {"Division": 0, "Sub Division": 1, "Office": 2}[level]
    notes, plan = [], []
    for kind, what, sort, desc, bm in (("delivery", "Actual Delivery", 4, not worst_first, bench[0]),
                                       ("redrt", "Return & Redirection", 11, worst_first, bench[1])):
        for sh in PPT_PRODUCT:
            if sh not in early or sh not in later:
                if kind == "delivery" and (sh in early or sh in later):
                    notes.append(f"{PPT_PRODUCT[sh]}: files are needed for both periods - no slides.")
                continue
            x, y = early[sh], later[sh]
            if need > min(x[3], y[3]):
                if kind == "delivery":
                    notes.append(f"{PPT_PRODUCT[sh]}: {level}-wise slides need "
                                 f"{'an Office Level' if level == 'Office' else 'a Sub Division Level or Office Level'} export.")
                continue
            if bool(x[2]) != bool(y[2]) and kind == "delivery":
                notes.append(f"{PPT_PRODUCT[sh]}: BOs file given for one period only - only the All Offices slides.")
            s1, s2 = dict(hq_scopes(x[1], x[2], level)), dict(hq_scopes(y[1], y[2], level))
            for sc in [k for k in s1 if k in s2]:
                nh, table = trio_table(level, s1[sc], s2[sc], False, sort, desc, 1)       # % on Invoiced, sorted
                data, total = table[:-1], table[-1]
                if not data:
                    continue
                cut = bool(max_rows) and len(data) > max_rows
                data = data[:max_rows] if cut else data
                pages = [data[i:i + rows_per] for i in range(0, len(data), rows_per)]
                plan.append((f"D+0 {PPT_PRODUCT[sh]} \u2013 {what} -{wording}", PPT_TAG[sc], bm, nh, pages, total, kind, cut))
    total_slides = sum(len(p[4]) for p in plan)
    if not total_slides:
        raise ValueError(" ".join(notes) or "No slides could be made.")
    if total_slides > max_slides:
        raise ValueError(f"That would be {total_slides} slides (limit {max_slides}). Use Sub Division-wise, or set "
                         "'Max rows per table', or a larger 'Rows per slide'.")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for title, tag, bm, nh, pages, total, kind, cut in plan:
        for i, rows in enumerate(pages):
            t = title + (f" ({i + 1}/{len(pages)})" if len(pages) > 1 else "")
            _ppt_slide(prs, t, tag, bm, ranges, level, nh, rows, total if i == len(pages) - 1 else None, kind, cut)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue(), notes, total_slides


# ---------- 2) Compare two report files ----------
def parse_block(raw, hdr, name_col):
    roles = {}
    for j in range(name_col + 1, raw.shape[1]):
        h = raw.iat[hdr, j]
        t = norm(h) if isinstance(h, str) else ""
        if not t or "%" in t or "(" in t:
            continue
        role = ("rec" if t.startswith("received") else "inv" if t.startswith("invoiced") else
                "dl" if "deliver" in t else "rd" if "redirect" in t else "rt" if "return" in t else None)
        if role:
            roles.setdefault(role, j)
    if all(k in roles for k in ("rec", "inv", "dl", "rd", "rt")):
        cols = [roles[k] for k in ("rec", "inv", "dl", "rd", "rt")]
    elif raw.shape[1] >= name_col + 6:
        cols = list(range(name_col + 1, name_col + 6))      # fallback: fixed positions
    else:
        return None
    title = None
    for k in range(1, 4):
        if hdr - k >= 0:
            title = next((x.strip() for x in raw.iloc[hdr - k] if isinstance(x, str) and x.strip()), None)
        if title:
            break
    rows, i = {}, hdr + 1
    while i < len(raw):
        n = raw.iat[i, name_col]
        if not isinstance(n, str) or not n.strip():
            if rows:
                break
        elif not n.strip().lower().startswith("total"):
            rows[" ".join(n.lower().split())] = (" ".join(n.split()), [to_num(raw.iat[i, c]) for c in cols])
        i += 1
    return title, rows, i


def read_blocks(xl, sheet):
    """All report tables on a sheet -> [(title, {key: (division, [5 counts])})]."""
    raw = xl.parse(sheet, header=None)
    blocks, i = [], 0
    while i < len(raw):
        col = next((j for j in range(min(4, raw.shape[1]))
                    if isinstance(raw.iat[i, j], str) and
                    (norm(raw.iat[i, j]).startswith("nameof") or norm(raw.iat[i, j]) in ("division", "region"))), None)
        res = parse_block(raw, i, col) if col is not None else None
        if res:
            blocks.append((res[0] or sheet, res[1]))
            i = res[2]
        else:
            i += 1
    return blocks


def compare_rows(r1, r2):
    keys = list(r2) + [k for k in r1 if k not in r2]
    out, t1, t2 = [], [0.0] * 5, [0.0] * 5
    for k in keys:
        name = (r2.get(k) or r1.get(k))[0]
        c1 = r1[k][1] if k in r1 else [0.0] * 5
        c2 = r2[k][1] if k in r2 else [0.0] * 5
        t1 = [a + b for a, b in zip(t1, c1)]
        t2 = [a + b for a, b in zip(t2, c2)]
        out.append((name, c1, c2))
    out.append(("Total", t1, t2))
    res = []
    for name, c1, c2 in out:
        m1, m2 = metrics(c1), metrics(c2)
        res.append((name, m1, m2, [m2[0] - m1[0]] + [m2[i] - m1[i] for i in (2, 4, 6, 8, 10, 11)]))
    return res


def comparison_title(t1, t2):
    m1, m2 = PERIOD.search(t1), PERIOD.search(t2)
    if m1 and m2:
        return f"Comparison {t2[:m2.start()].rstrip(' -')} -{m1.group(0)} vs {m2.group(0)}"
    return f"Comparison: {t1} vs {t2}"


def run_comparison(bytes1, bytes2):
    x1, x2 = pd.ExcelFile(io.BytesIO(bytes1)), pd.ExcelFile(io.BytesIO(bytes2))
    wb = Workbook()
    wb.remove(wb.active)
    previews, notes = {}, []
    for sh in x1.sheet_names:
        if sh not in x2.sheet_names:
            notes.append(f"Sheet '{sh}' is not in the later file - skipped.")
            continue
        a, b = read_blocks(x1, sh), read_blocks(x2, sh)
        if not a or not b:
            notes.append(f"Sheet '{sh}' does not look like a report sheet - skipped.")
            continue
        if len(a) != len(b):
            notes.append(f"Sheet '{sh}' has {len(a)} table(s) in the earlier file and {len(b)} in the later "
                         f"file - only the first {min(len(a), len(b))} are compared.")
        blocks = []
        for (t1, r1), (t2, r2) in zip(a, b):
            blocks.append((t1, t2, comparison_title(t1, t2), compare_rows(r1, r2)))
        build_comparison_sheet(wb.create_sheet(sh[:31]), blocks)
        previews[sh] = [(tc, pd.DataFrame([[r[0]] + r[3] for r in rows], columns=[NAME] + COMP_HEAD))
                        for _, _, tc, rows in blocks]
    if not previews:
        raise ValueError("No matching report sheets found in the two files.")
    return to_bytes(wb), previews, notes


# ---------- Streamlit screens ----------
def df_show(holder, obj):
    """Full-width table; works on old and new Streamlit versions."""
    data = obj.data if obj.__class__.__name__ == "Styler" else obj
    for c in data.columns:
        if str(c).lower().replace(".", "") == "sno":          # S.No. / S.no: text, so a blank Total row is fine
            data[c] = data[c].astype(str)
    try:
        holder.dataframe(obj, hide_index=True, width="stretch")
    except TypeError:
        holder.dataframe(obj, hide_index=True, use_container_width=True)


def login(st):
    try:
        users = dict(st.secrets["users"])
    except Exception:
        return True                      # no users configured -> no login screen
    if st.session_state.get("auth"):
        return True
    st.title("Login")
    with st.form("login"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Submit"):
            if u in users and hmac.compare_digest(str(users[u]), p):
                st.session_state["auth"] = True
                st.rerun()
            st.error("Wrong username or password")
    return False


def page_generate(st):
    sb = st.sidebar
    sb.divider()
    sb.header("Report Period")
    c1, c2 = sb.columns(2)
    today = datetime.date.today()
    d1 = c1.date_input("From", today.replace(day=1))
    d2 = c2.date_input("To", today)
    sb.divider()
    sb.header("Upload Report Files")
    up = lambda label, key: sb.file_uploader(label, type=["csv", "xlsx"], key=key)
    sb.subheader("1️⃣ All Products")
    ap_all, ap_bo = up("All Offices file", "ap_all"), up("BOs file (optional)", "ap_bo")
    sb.subheader("2️⃣ Speed Post Document (Speed Letter)")
    sl_all, sl_bo = up("All Offices file", "sl_all"), up("BOs file (optional)", "sl_bo")
    sb.subheader("3️⃣ Consolidated Parcels (Speed Parcel + Regd. Parcel)")
    sp_all, rp_all = up("Speed Parcel - All Offices", "sp_all"), up("Regd. Parcel - All Offices", "rp_all")
    sp_bo, rp_bo = up("Speed Parcel - BOs (optional)", "sp_bo"), up("Regd. Parcel - BOs (optional)", "rp_bo")

    st.title("Delivery Transit Analysis (D+0)")
    st.info("Upload the raw files from the sidebar to generate the D+0 reports.\n\n"
            "1) **All Products** - All Offices file, and optionally the BOs file (also gives BOs and Excluding-BOs).\n\n"
            "2) **Speed Post Document** - same pattern.\n\n"
            "3) **Consolidated Parcels** - Speed Parcel + Regd. Parcel are added together; "
            "Excluding-BOs = All Offices - BOs.")
    groups = [("All Products", "All Products", [ap_all], [ap_bo], 1),
              ("Speed Letter", "Speed Post Document", [sl_all], [sl_bo], 1),
              ("Parcels", "Consolidated Parcels", [sp_all, rp_all], [sp_bo, rp_bo], 2)]
    jobs = []
    try:
        for sheet, rname, alls, bos, need in groups:
            alls, bos = [f for f in alls if f], [f for f in bos if f]
            if not alls:
                continue
            if len(alls) < need or (bos and len(bos) < need):
                st.warning(f"{rname}: only some of the files were uploaded, so the figures are incomplete.")
            jobs.append((sheet, rname, [read_raw(f) for f in alls], [read_raw(f) for f in bos]))
    except Exception as e:
        st.error(f"Could not read the files: {e}")
        return
    if not jobs:
        return
    data, previews = make_reports(jobs, d1, d2)
    st.download_button("Download All Reports (Master Excel Workbook)", data,
                       "D0_Delivery_Reports.xlsx", XLSX, type="primary")
    fmt = {h: "{:.2%}" if i in PCT else "{:,.0f}" for i, h in enumerate(HEAD)}
    for tab, blocks in zip(st.tabs(list(previews)), previews.values()):
        for title, df in blocks:
            tab.subheader(title)
            df_show(tab, df.style.format(fmt))


def page_ndc_dc(st):
    sb = st.sidebar
    sb.divider()
    sb.header("Report Period")
    c1, c2 = sb.columns(2)
    today = datetime.date.today()
    d1 = c1.date_input("From", today.replace(day=1), key="nd_d1")
    d2 = c2.date_input("To", today, key="nd_d2")
    sb.divider()
    sb.header("Upload Report Files")
    f_ndc = sb.file_uploader("NDCs file", type=["csv", "xlsx"], key="ndc_file")
    f_dc = sb.file_uploader("DCs file", type=["csv", "xlsx"], key="dc_file")
    st.title("Delivery Transit Analysis (D+0) - NDCs / DCs")
    st.info("Upload the **NDCs** file and/or the **DCs** file from the sidebar. "
            "Each gets its own office-wise D+0 report, and both come in one master Excel.")
    jobs = []
    try:
        for sheet, f in [("NDCs", f_ndc), ("DCs", f_dc)]:
            if not f:
                continue
            offices, types = read_offices(f)
            if sheet == "NDCs" and types and "NDC" not in types:
                st.warning(f"{f.name} does not look like an NDC file (Office Type: {', '.join(sorted(types))}).")
            if sheet == "DCs" and "NDC" in types:
                st.warning(f"{f.name} looks like an NDC file - did you put it in the DCs box?")
            jobs.append((sheet, sheet, offices))
    except Exception as e:
        st.error(f"Could not read the files: {e}")
        return
    if not jobs:
        return
    data, previews = make_office_reports(jobs, d1, d2)
    st.download_button("Download NDC / DC Reports (Excel)", data, "NDC_DC_D0_Reports.xlsx", XLSX, type="primary")
    fmt = {h: "{:.2%}" if i in PCT else "{:,.0f}" for i, h in enumerate(HEAD)}
    for tab, blocks in zip(st.tabs(list(previews)), previews.values()):
        for title, df in blocks:
            tab.subheader(title)
            df_show(tab, df.style.format(fmt))


def page_day(st):
    sb = st.sidebar
    sb.divider()
    sb.header("Dates")
    today = datetime.date.today()
    d1 = sb.date_input("Earlier date", today - datetime.timedelta(days=2), key="dd_d1")
    d2 = sb.date_input("Later date", today - datetime.timedelta(days=1), key="dd_d2")
    sb.divider()
    sb.header("Upload Report Files")
    f1 = sb.file_uploader("Earlier date file", type=["csv", "xlsx"], key="dd_f1")
    f2 = sb.file_uploader("Later date file", type=["csv", "xlsx"], key="dd_f2")
    sb.divider()
    sb.header("Offices and Format")
    grp = sb.selectbox("Offices to include", list(GROUPS) + [LIST_GROUP], key="dd_grp")
    is_list, names = grp == LIST_GROUP, OFFICES_26
    if is_list:
        text = sb.text_area("Offices (one per line - edit to change the list)", "\n".join(OFFICES_26),
                            height=320, key="dd_list")
        names = [x.strip() for x in text.splitlines() if x.strip()]
    fmt = sb.selectbox("Report format", FORMATS, index=1 if is_list else 0)       # follows the office choice
    msort = sb.selectbox("Sort rows by", MISSENT_SORT, index=0 if is_list else 1) if fmt == FORMATS[1] else None
    title = sb.text_input("Report title", "Speed Post Document", key="dd_title")
    st.title("Day-wise Comparison - Redirection / Missent")
    st.info("Upload the full all-offices file of the **earlier date** and of the **later date** (same product, "
            "e.g. Speed Letter). Choose the offices (DCs, NDCs, NDCs + DCs, or the list of 26 offices) and the "
            "format: **Redirection** (% on Received, Summary + Detailed sheets, sorted by % Inc/Dec) or "
            "**Missent** (% on Same Day Invoiced, one sheet).")
    if not (f1 and f2):
        return
    try:
        o1, o2 = read_offices(f1)[0], read_offices(f2)[0]
        if is_list:
            (g1, m1), (g2, m2) = pick_list(o1, names), pick_list(o2, names)
            if not names or len(m1) == len(g1) or len(m2) == len(g2):
                raise ValueError("none of the listed offices were found in one of the files")
        else:
            g1, g2, m1, m2 = pick_group(o1, GROUPS[grp]), pick_group(o2, GROUPS[grp]), [], []
            if not g1 or not g2:
                raise ValueError("no matching offices found in one of the files")
        if fmt == FORMATS[0]:
            rows, total = day_compare(g1, g2)
        else:
            mrows, mtotal = missent_rows(g1, g2, msort)
    except Exception as e:
        st.error(f"Could not make the comparison: {e}")
        return
    for lab, miss in (("earlier", m1), ("later", m2)):
        if miss:
            st.warning(f"Not found in the {lab} date file (counted as zero): " + ", ".join(miss))
    if not is_list:
        for a, b, lab in ((g1, g2, "later"), (g2, g1, "earlier")):
            miss = [a[k][0] for k in a if k not in b]
            if miss:
                st.warning(f"Not in the {lab} date file (counted as zero there): " + ", ".join(miss))
    if fmt == FORMATS[1]:
        lab1, lab2 = f"{d1:%d.%m.%Y}", f"{d2:%d.%m.%Y}"
        st.caption(f"{len(mrows)} offices compared.")
        st.download_button("Download Missent Comparison (Excel)",
                           missent_workbook(f"Missent Comparison for {title}", lab1, lab2, mrows, mtotal),
                           "Missent_Comparison.xlsx", XLSX, type="primary")
        cols = (["S.no", "Office Name"] + [f"{lab} - {h}" for lab in (lab1, lab2)
                                          for h in ("Same Day Invoiced", "D0 Redirected", "%Missent")]
                + ["Increase/Decrease"])
        table = [[i + 1, r[0]] + r[1] for i, r in enumerate(mrows)] + [["", "Total"] + mtotal[1]]
        sty = pd.DataFrame(table, columns=cols).style.format(
            {c: "{:.2%}" if ("%" in c or c == "Increase/Decrease") else "{:,.0f}" for c in cols[2:]})
        mapper = getattr(sty, "map", None) or sty.applymap
        sty = mapper(lambda v: "color: red" if v > 1e-9 else "color: green" if v < -1e-9 else "",
                     subset=["Increase/Decrease"])
        df_show(st, sty)
        return
    st.caption(f"{len(rows)} offices compared.")
    st.download_button("Download Comparison Excel", day_workbook(title, d1, d2, rows, total),
                       "Day_wise_Comparison.xlsx", XLSX, type="primary")
    for tab, detailed in zip(st.tabs(["Summary", "Detailed"]), (False, True)):
        heads, table = day_table(rows, total, detailed)
        cols = (["S.No.", "Division Name", "Office Name"] + [f"{d1:%d.%m} {h}" for h in heads]
                + [f"{d2:%d.%m} {h}" for h in heads] + ["% Inc/Dec"])
        fmt_ = {c: "{:.2%}" if ("%" in c) else "{:,.0f}" for c in cols[3:]}
        df_show(tab, pd.DataFrame(table, columns=cols).style.format(fmt_))


def hq_uploaders(sb, prefix, label, expanded=False):
    """The three upload sections (All Products / Speed Letter / Parcels) inside an expander."""
    ex = sb.expander(label, expanded=expanded)
    up = lambda lab, key: ex.file_uploader(lab, type=["csv", "xlsx"], key=f"{prefix}_{key}")
    ex.subheader("1️⃣ All Products")
    f = {"ap_all": up("All Offices file", "ap_all"), "ap_bo": up("BOs file (optional)", "ap_bo")}
    ex.subheader("2️⃣ Speed Post Document (Speed Letter)")
    f.update(sl_all=up("All Offices file", "sl_all"), sl_bo=up("BOs file (optional)", "sl_bo"))
    ex.subheader("3️⃣ Consolidated Parcels (Speed Parcel + Regd. Parcel)")
    f.update(sp_all=up("Speed Parcel - All Offices", "sp_all"), rp_all=up("Regd. Parcel - All Offices", "rp_all"),
             sp_bo=up("Speed Parcel - BOs (optional)", "sp_bo"), rp_bo=up("Regd. Parcel - BOs (optional)", "rp_bo"))
    return f


def hq_read_period(st, f, tag):
    """Uploaded files of one period -> {sheet: (report name, [All rows], [BO rows], level rank)}."""
    out = {}
    for sheet, rname, ak, bk, need in HQ_GROUPS:
        alls, bos = [f[k] for k in ak if f[k]], [f[k] for k in bk if f[k]]
        if not alls:
            continue
        if len(alls) < need or (bos and len(bos) < need):
            st.warning(f"{tag} - {rname}: only some of the files were uploaded, so the figures are incomplete.")
        ra, rb = [read_hq(x) for x in alls], [read_hq(x) for x in bos]
        out[sheet] = (rname, [x[0] for x in ra], [x[0] for x in rb], min(x[1] for x in ra + rb))
    return out


def page_hq(st):
    sb = st.sidebar
    today = datetime.date.today()
    last_prev = today.replace(day=1) - datetime.timedelta(days=1)
    sb.divider()
    dates = {}
    for k, head, cap, d_from, d_to in [
            (1, "Period 1", "The base period. Period 2 (+ Period 3) is compared with it.", last_prev.replace(day=1), last_prev),
            (2, "Period 2", "", today.replace(day=1), today),
            (3, "Period 3 (optional)", "If given, Period 2 + Period 3 are added together and compared with Period 1.", today, today)]:
        sb.header(head)
        if cap:
            sb.caption(cap)
        ca, cb = sb.columns(2)
        dates[k] = (ca.date_input("From", d_from, key=f"hq_a{k}"), cb.date_input("To", d_to, key=f"hq_b{k}"))
    sb.divider()
    sb.header("Report Options")
    base = sb.radio("Percentages based on", ["Received", "Invoiced", "Both (two sets of sheets)"], key="hq_base")
    lv = sb.multiselect("Report level", list(LEVELS), default=["Division-wise"], key="hq_lv")
    sort_by = sb.selectbox("Sort by", list(SORT_OPTIONS), list(SORT_OPTIONS).index(SORT_DEFAULT), key="hq_sort")
    order = sb.radio("Order", ORDERS, key="hq_order")
    sort_on = sb.radio("Sort using (when comparing)", ["Earlier period", "Later period", "Comparison"], index=1,
                       key="hq_on")
    sb.divider()
    sb.header("Upload Report Files")
    sb.caption("Upload the export from the portal at any level. **Office Level** gives Division, Sub Division "
               "and Office wise reports; **Sub Division Level** gives Division and Sub Division wise; "
               "**Division Level** gives Division wise. Use the same level for all periods.")
    files = {1: hq_uploaders(sb, "hq1", "Period 1 files", True),
             2: hq_uploaders(sb, "hq2", "Period 2 files"),
             3: hq_uploaders(sb, "hq3", "Period 3 files (optional)")}

    st.title("HQ Delivery Transit Analysis (D+0)")
    st.info("Upload **Period 1** files to get the D+0 report. Add **Period 2** and it is compared with Period 1. "
            "Add **Period 3** as well and Period 2 + Period 3 are added together, and that total is compared "
            "with Period 1. The Excel has, for each of All Products, Speed Letter and Parcels: every period's "
            "report, Period 2 vs Period 3, Period 2 + Period 3 (addition only), and the comparison with Period 1. "
            "Comparison columns show later minus earlier (green = better, red = worse for the % columns).")
    if not lv:
        st.warning("Select at least one report level in the sidebar.")
        return
    try:
        P = {k: hq_read_period(st, files[k], f"Period {k}") for k in (1, 2, 3)}
    except Exception as e:
        st.error(f"Could not read the files: {e}")
        return
    if not any(P.values()):
        return
    bases = {"Received": [True], "Invoiced": [False]}.get(base, [True, False])
    try:
        data, previews, notes = make_hq_workbook(P, dates, [LEVELS[k] for k in LEVELS if k in lv], bases,
                                                 SORT_OPTIONS[sort_by], order == ORDERS[1],
                                                 ["Earlier period", "Later period", "Comparison"].index(sort_on))
    except Exception as e:
        st.error(f"Could not make the report: {e}")
        return
    for n in notes:
        st.warning(n)
    multi = sum(1 for k in P if P[k]) > 1
    st.download_button("Download HQ D+0 Reports (Master Excel Workbook)", data,
                       "HQ_D0_Multi_Period_Reports.xlsx" if multi else "HQ_D0_Reports.xlsx", XLSX, type="primary")
    st.caption(f"{len(previews)} sheets in the Excel.")
    pairs = hq_ppt_pairs(P, dates)
    if pairs and st.checkbox("Also create PowerPoint slides (Delivery + Return & Redirection)", key="hq_ppt"):
        c1, c2, c3 = st.columns(3)
        pick_cmp = c1.selectbox("Slides compare", list(pairs), key="hq_ppt_cmp")
        ppt_level = c2.selectbox("Slides level", list(PPT_LEVELS), key="hq_ppt_lv")
        words = c3.text_input("Title wording", "Week on Week Comparison", key="hq_ppt_words")
        c4, c5, c6 = st.columns(3)
        bm_d = c4.text_input("Delivery benchmark % (optional)", "", key="hq_ppt_bd")
        bm_r = c5.text_input("Return + Redirection benchmark % (optional)", "", key="hq_ppt_br")
        order = c6.selectbox("Row order", ["Best first (as in the sample)", "Worst first"], key="hq_ppt_ord")
        c7, c8, _ = st.columns(3)
        rows_per = int(c7.number_input("Rows per slide", min_value=5, max_value=30, value=18, key="hq_ppt_rows"))
        max_rows = int(c8.number_input("Max rows per table (0 = all)", min_value=0, max_value=5000, value=0,
                                       key="hq_ppt_max"))
        fsig = tuple((k, name, getattr(f, "name", ""), getattr(f, "size", 0))
                     for k in (1, 2, 3) for name, f in files[k].items() if f)
        sig = repr((pick_cmp, ppt_level, words, bm_d, bm_r, order, rows_per, max_rows, fsig, sorted(dates.items())))
        if st.button("Create PowerPoint slides", key="hq_ppt_btn"):
            try:
                early, later, ra, rb = pairs[pick_cmp]
                with st.spinner("Creating the slides..."):
                    ppt, pnotes, nslides = make_hq_ppt(early, later, (ra, rb), PPT_LEVELS[ppt_level], words,
                                                       (bench_val(bm_d), bench_val(bm_r)), rows_per, max_rows,
                                                       order.startswith("Worst"))
                st.session_state["hq_ppt_res"] = (sig, ppt, pnotes, nslides)
            except ImportError:
                st.error("PowerPoint needs the python-pptx package: add a line 'python-pptx' to requirements.txt on "
                         "GitHub, then reboot the app.")
            except Exception as e:
                st.error(f"Could not make the slides: {e}")
        res = st.session_state.get("hq_ppt_res")
        if res and res[0] == sig:
            for n in res[2]:
                st.warning(n)
            st.download_button(f"Download PowerPoint slides ({res[3]} slides)", res[1], "HQ_D0_Slides.pptx",
                               "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        elif res:
            st.info("Settings or files changed - press 'Create PowerPoint slides' again.")
    if len(previews) > 6:
        pick = st.selectbox("Preview sheet", list(previews))
        views = [(st.container(), previews[pick])]
    else:
        views = list(zip(st.tabs(list(previews)), previews.values()))
    for holder, blocks in views:
        for title, df, fm in blocks:
            holder.subheader(title)
            fmt = {c: FMT_PY[f] for c, f in zip(df.columns, fm) if f}
            df_show(holder, df.style.format(fmt))


def page_text(st):
    sb = st.sidebar
    sb.divider()
    sb.header("Upload Report File")
    f = sb.file_uploader("Daily report Excel (with the 'Morning report' and '24 Morning report' sheets)",
                         type=["xlsx", "xlsm"], key="tx_file")
    sb.date_input("Report date (used only if the sheet has no date)", datetime.date.today(), key="tx_date")
    sb.divider()
    sb.header("Options")
    num = lambda label, v, key: sb.number_input(label, min_value=0.0, max_value=100.0, value=v, step=0.5,
                                                format="%.1f", key=key)
    opts = {}
    for kind, head, d in (("ndc", "NDCs", (95.0, 85.0, 5.0, 5.0)), ("dc", "DCs", (95.0, 85.0, 5.0, 15.0)),
                          ("24", "24 Speed Post (NDCs/DCs)", (85.0, 85.0, 5.0, 10.0))):
        sb.subheader(head)
        opts[kind] = (num("Daily target %", d[0], f"tx_{kind}_t"),
                      num("Successful delivery mark % (list offices below)", d[1], f"tx_{kind}_m"),
                      num("Redirection above %", d[2], f"tx_{kind}_r"),
                      num("Return above %", d[3], f"tx_{kind}_x"))
    st.title("Text for DC / NDC / 24 Reports")
    st.info("Upload the daily Excel. The texts are made from the **Morning report** sheet (NDCs and DCs) and the "
            "**24 Morning report** sheet. Set the successful delivery mark and the Redirection / Return limits "
            "in the sidebar, then use the copy button on each text.")
    if not f:
        return
    try:
        secs = read_text_workbook(f.getvalue())
    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")
        return
    fallback, texts = st.session_state.get("tx_date", datetime.date.today()), []
    for kind, tab_name in (("ndc", "NDCs"), ("dc", "DCs"), ("24", "24 Speed Post")):
        if kind not in secs:
            st.warning(f"{tab_name}: table not found in the sheet - skipped.")
            continue
        try:
            texts.append((tab_name, make_text(kind, secs[kind], fallback, *opts[kind])))
        except Exception as e:
            st.warning(f"{tab_name}: {e}")
    if not texts:
        return
    st.download_button("Download all texts (.txt)", ("\n\n" + "-" * 30 + "\n\n").join(t for _, t in texts),
                       "DC_NDC_24_texts.txt", "text/plain")
    for tab, (_, text) in zip(st.tabs([n for n, _ in texts]), texts):
        try:
            tab.code(text, language=None, wrap_lines=True)
        except TypeError:
            tab.code(text, language=None)


def page_dss(st):
    sb = st.sidebar
    today = datetime.date.today()
    last_prev = today.replace(day=1) - datetime.timedelta(days=1)
    sb.divider()
    sb.header("Report Options")
    lv = sb.multiselect("Report level", list(LEVELS), default=["Division-wise"], key="dss_lv")
    sort_by = sb.selectbox("Sort by", DSS_SORT, key="dss_sort")
    order = sb.radio("Order", ORDERS, key="dss_order")
    hide_zero = sb.checkbox("Hide offices with no articles in both periods", value=True, key="dss_hide")
    files, dates = {}, {}
    for k, d_from, d_to in ((1, last_prev.replace(day=1), last_prev), (2, today.replace(day=1), today)):
        sb.divider()
        sb.header(f"Period {k}")
        ca, cb = sb.columns(2)
        dates[k] = (ca.date_input("From", d_from, key=f"dss_a{k}"), cb.date_input("To", d_to, key=f"dss_b{k}"))
        files[k] = sb.file_uploader(f"Period {k} file (Delivery Mobile App Usage export)", type=["csv", "xlsx"],
                                    key=f"dss_f{k}")
    st.title("DSS App Report")
    st.info("Pick the report level and the dates of the two periods in the sidebar, then upload the "
            "**Delivery Mobile App Usage** export of each period. Mobile Usage = DSS article count / PDM article "
            "count, and % Difference = Period 2 usage minus Period 1 usage (green = usage went up, red = went down).")
    if not lv:
        st.warning("Select at least one report level in the sidebar.")
        return
    if not (files[1] and files[2]):
        return
    try:
        (r1, k1), (r2, k2) = read_dss(files[1]), read_dss(files[2])
        txt = lambda k: f"{dates[k][0]:%d.%m.%Y} - {dates[k][1]:%d.%m.%Y}"
        data, previews, notes = make_dss_report(r1, k1, r2, k2, [LEVELS[k] for k in LEVELS if k in lv],
                                                txt(1), txt(2), hide_zero, sort_by, order == ORDERS[1])
    except Exception as e:
        st.error(f"Could not make the report: {e}")
        return
    for n in notes:
        st.warning(n)
    st.download_button("Download DSS App Report (Excel)", data, "DSS_App_Report.xlsx", XLSX, type="primary")
    for tab, df in zip(st.tabs([f"{k}-wise" for k in previews]), previews.values()):
        fmt = {c: ("{:.2%}" if ("Usage" in c or "Difference" in c) else "{:,.0f}") for c in df.columns[1 + df.shape[1] - 7 - 1:]
               if c != df.columns[0]}
        sty = df.style.format({c: f for c, f in fmt.items() if c in df.columns})
        diff = [c for c in df.columns if c == "% Difference"]
        mapper = getattr(sty, "map", None) or sty.applymap
        sty = mapper(lambda v: "color: red" if v < -1e-9 else "color: green" if v > 1e-9 else "", subset=diff)
        df_show(tab, sty)


def page_cod(st):
    sb = st.sidebar
    today = datetime.date.today()
    sb.divider()
    sb.header("Report Options")
    lv = sb.multiselect("Report level", list(LEVELS), default=["Division-wise"], key="cod_lv")
    sort_by = sb.selectbox("Sort by (latest period)", COD_SORT, key="cod_sort")
    order = sb.radio("Order", ORDERS, index=1, key="cod_order")
    files, dates = {}, {}
    for k, d_from, d_to in ((1, today - datetime.timedelta(days=14), today - datetime.timedelta(days=8)),
                            (2, today - datetime.timedelta(days=7), today - datetime.timedelta(days=1))):
        sb.divider()
        sb.header(f"Period {k}" + (" (optional)" if k == 2 else ""))
        ca, cb = sb.columns(2)
        dates[k] = (ca.date_input("Delivery date from", d_from, key=f"cod_a{k}"),
                    cb.date_input("Delivery date to", d_to, key=f"cod_b{k}"))
        files[k] = sb.file_uploader(f"Period {k} file (COD Collection export)", type=["csv", "xlsx"], key=f"cod_f{k}")
    st.title("COD Collection Report")
    st.info("Upload the **COD Collection Performance** export of one period to get its report, or of both periods "
            "to get Period 1, Period 2 and their comparison (3 sheets). % of Digital Transactions = Digital count / "
            "Total COD count, and Difference = Period 2 minus Period 1 (percentage points).")
    if not lv:
        st.warning("Select at least one report level in the sidebar.")
        return
    if not (files[1] or files[2]):
        return
    try:
        txt = lambda k: f"{dates[k][0]:%d.%m.%Y} to {dates[k][1]:%d.%m.%Y}"
        ps = []
        for k in (1, 2):
            if files[k]:
                rows, rank = read_cod(files[k])
                ps.append((rows, rank, txt(k)))
            else:
                ps.append(None)
        data, previews, notes = make_cod_report(ps[0], ps[1], [LEVELS[k] for k in LEVELS if k in lv], sort_by,
                                                order == ORDERS[1])
    except Exception as e:
        st.error(f"Could not make the report: {e}")
        return
    for n in notes:
        st.warning(n)
    st.download_button("Download COD Collection Report (Excel)", data, "COD_Collection_Report.xlsx", XLSX,
                       type="primary")
    if len(previews) > 3:
        pick = st.selectbox("Preview sheet", list(previews))
        views = [(st.container(), previews[pick])]
    else:
        views = list(zip(st.tabs(list(previews)), previews.values()))
    for holder, (title, df) in views:
        holder.subheader(title)
        fmt = {c: "{:.2f}" if ("%" in c or c.endswith("Difference")) else "{:,.0f}" for c in df.columns
               if any(h in c for h in COD_HEADS) or c == "Difference"}
        df_show(holder, df.style.format(fmt, na_rep=""))


def page_compare(st):
    st.title("Delivery Transit Analysis (D+0) - Period Comparison")
    st.info("Upload the two report Excel files you want to compare (reports made by this app, or the "
            "master workbooks). The comparison is built for every sheet that exists in both files.")
    c1, c2 = st.columns(2)
    f1 = c1.file_uploader("1) Earlier period report", type=["xlsx", "xlsm"])
    f2 = c2.file_uploader("2) Later period report", type=["xlsx", "xlsm"])
    if not (f1 and f2):
        return
    try:
        data, previews, notes = run_comparison(f1.getvalue(), f2.getvalue())
    except Exception as e:
        st.error(f"Could not compare these files: {e}")
        return
    for n in notes:
        st.warning(n)
    st.download_button("Download comparison Excel", data, "Comparison_Report.xlsx", XLSX, type="primary")
    fmt = {COMP_HEAD[0]: "{:+,.0f}", **{c: "{:+.2%}" for c in COMP_HEAD[1:]}}
    for tab, blocks in zip(st.tabs(list(previews)), previews.values()):
        for title, df in blocks:
            tab.subheader(title)
            df_show(tab, df.style.format(fmt))


def main():
    import streamlit as st
    st.set_page_config(page_title="D+0 Analytics", layout="wide")
    if not login(st):
        return
    pages = {"Region wise D+0 Reports": page_generate, "HQ D+0 Reports": page_hq,
             "Generate NDC / DC Reports": page_ndc_dc,
             "DCs Redirection Comparison": page_day, "Compare Two Periods": page_compare,
             "Text for DC/NDC/24 Reports": page_text, "DSS App report": page_dss,
             "COD Collection Report": page_cod}
    pages[st.sidebar.radio("Select", list(pages))](st)


if __name__ == "__main__":
    main()

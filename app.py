"""Delivery Transit Analysis (D+0): 1) generate reports from raw CSVs, 2) compare two periods."""
import io, re, hmac, datetime
import pandas as pd
from openpyxl import Workbook
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
GROUPS = {"DCs (DC / IDC)": re.compile(r"\bI?DC\b", re.I), "NDCs": re.compile(r"\bNDC\b", re.I)}
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


# ---------- 1d) HQ D+0 reports (Division / Sub Division / Office level, from Office Level exports) ----------
LEVELS = {"Division-wise": "Division", "Sub Division-wise": "Sub Division", "Office-wise": "Office"}
LEVEL_PARENTS = {"Division": [], "Sub Division": ["Division"], "Office": ["Division", "Sub Division"]}
LEVEL_SHORT = {"Division": "Div", "Sub Division": "Sub Div", "Office": "Office"}
LEVEL_TITLE = {"Division": "", "Sub Division": "Sub Division wise - ", "Office": "Office wise - "}
HQ_FMT = ["#,##0", "#,##0", "0.00%", "#,##0", "0.00%", "#,##0", "0.00%", "#,##0", "0.00%", "0.00%", "#,##0", "0.00%"]


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


def hq_block(title, level, d, recv):
    """One table. recv=True -> all % on Received, False -> all % on Invoiced. Sorted by Return+Redirection %."""
    on = "On received" if recv else "On invoiced"
    den = (lambda c: c[0]) if recv else (lambda c: c[1])
    items = [v for v in d.values() if any(v[2])]
    items.sort(key=lambda v: sd(v[2][3] + v[2][4], den(v[2])))
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


def make_hq_reports(jobs, levels, bases, d1, d2):
    """jobs = [(sheet, report name, [All Offices row-lists], [BO row-lists], file level rank 0/1/2)];
    bases = [True (received), False (invoiced)]. Returns (bytes, previews, notes)."""
    wb = Workbook()
    wb.remove(wb.active)
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
                blocks = [hq_block(f"Delivery Transit Analysis {rname}- {LEVEL_TITLE[level]}{sc} "
                                   f"-{d1:%d.%m.%Y} to {d2:%d.%m.%Y}", level, d, recv) for sc, d in scopes]
                name = (sheet + (f" - {LEVEL_SHORT[level]}" if len(levels) > 1 else "")
                        + ((" (Rcvd)" if recv else " (Inv)") if len(bases) > 1 else ""))[:31]
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
    if not previews:
        raise ValueError(" ".join(notes) or "Nothing to make.")
    return to_bytes(wb), previews, notes


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
            tab.dataframe(df.style.format(fmt), hide_index=True, use_container_width=True)


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
            tab.dataframe(df.style.format(fmt), hide_index=True, use_container_width=True)


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
    grp = sb.selectbox("Offices to include", list(GROUPS))
    title = sb.text_input("Report title", "Speed Post Document")
    st.title("Day-wise Comparison - Redirection")
    st.info("Upload the full all-offices file of the **earlier date** and of the **later date** "
            "(same product, e.g. Speed Letter). The DCs are picked out automatically, compared, and "
            "sorted by % Inc/Dec (biggest improvement first).")
    if not (f1 and f2):
        return
    try:
        g1 = pick_group(read_offices(f1)[0], GROUPS[grp])
        g2 = pick_group(read_offices(f2)[0], GROUPS[grp])
        if not g1 or not g2:
            raise ValueError("no matching offices found in one of the files")
        rows, total = day_compare(g1, g2)
    except Exception as e:
        st.error(f"Could not make the comparison: {e}")
        return
    for a, b, lab in ((g1, g2, "later"), (g2, g1, "earlier")):
        for k in a:
            if k not in b:
                st.warning(f"{a[k][0]} is missing in the {lab} date file - counted as zero there.")
    st.caption(f"{len(rows)} offices compared.")
    st.download_button("Download Comparison Excel", day_workbook(title, d1, d2, rows, total),
                       "Day_wise_Comparison.xlsx", XLSX, type="primary")
    for tab, detailed in zip(st.tabs(["Summary", "Detailed"]), (False, True)):
        heads, table = day_table(rows, total, detailed)
        cols = (["S.No.", "Division Name", "Office Name"] + [f"{d1:%d.%m} {h}" for h in heads]
                + [f"{d2:%d.%m} {h}" for h in heads] + ["% Inc/Dec"])
        fmt = {c: "{:.2%}" if ("%" in c) else "{:,.0f}" for c in cols[3:]}
        tab.dataframe(pd.DataFrame(table, columns=cols).style.format(fmt), hide_index=True, use_container_width=True)


def page_hq(st):
    sb = st.sidebar
    sb.divider()
    sb.header("Report Period")
    c1, c2 = sb.columns(2)
    today = datetime.date.today()
    d1 = c1.date_input("From", today.replace(day=1), key="hq_d1")
    d2 = c2.date_input("To", today, key="hq_d2")
    sb.divider()
    sb.header("Report Options")
    base = sb.radio("Percentages based on", ["Received", "Invoiced", "Both (two sets of sheets)"], key="hq_base")
    lv = sb.multiselect("Report level", list(LEVELS), default=["Division-wise"], key="hq_lv")
    sb.divider()
    sb.header("Upload Report Files")
    sb.caption("Upload the export from the portal at any level. **Office Level** gives Division, Sub Division "
               "and Office wise reports; **Sub Division Level** gives Division and Sub Division wise; "
               "**Division Level** gives Division wise.")
    up = lambda label, key: sb.file_uploader(label, type=["csv", "xlsx"], key=key)
    sb.subheader("1️⃣ All Products")
    ap_all, ap_bo = up("All Offices file", "hq_ap_all"), up("BOs file (optional)", "hq_ap_bo")
    sb.subheader("2️⃣ Speed Post Document (Speed Letter)")
    sl_all, sl_bo = up("All Offices file", "hq_sl_all"), up("BOs file (optional)", "hq_sl_bo")
    sb.subheader("3️⃣ Consolidated Parcels (Speed Parcel + Regd. Parcel)")
    sp_all, rp_all = up("Speed Parcel - All Offices", "hq_sp_all"), up("Regd. Parcel - All Offices", "hq_rp_all")
    sp_bo, rp_bo = up("Speed Parcel - BOs (optional)", "hq_sp_bo"), up("Regd. Parcel - BOs (optional)", "hq_rp_bo")

    st.title("HQ Delivery Transit Analysis (D+0)")
    st.info("Upload the raw files from the sidebar. Choose whether all % are worked out on "
            "**Received** or on **Invoiced** (or get both), and which levels you want "
            "(Division / Sub Division / Office). With a BOs file you also get the BOs and Excluding-BOs tables; "
            "Excluding-BOs = All Offices - BOs.")
    if not lv:
        st.warning("Select at least one report level in the sidebar.")
        return
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
            ra, rb = [read_hq(f) for f in alls], [read_hq(f) for f in bos]
            jobs.append((sheet, rname, [x[0] for x in ra], [x[0] for x in rb], min(x[1] for x in ra + rb)))
    except Exception as e:
        st.error(f"Could not read the files: {e}")
        return
    if not jobs:
        return
    bases = {"Received": [True], "Invoiced": [False]}.get(base, [True, False])
    try:
        data, previews, notes = make_hq_reports(jobs, [LEVELS[k] for k in LEVELS if k in lv], bases, d1, d2)
    except Exception as e:
        st.error(f"Could not make the report: {e}")
        return
    for n in notes:
        st.warning(n)
    st.download_button("Download HQ D+0 Reports (Master Excel Workbook)", data, "HQ_D0_Reports.xlsx", XLSX,
                       type="primary")
    for tab, blocks in zip(st.tabs(list(previews)), previews.values()):
        for title, df, fm in blocks:
            tab.subheader(title)
            fmt = {c: "{:.2%}" if f == "0.00%" else "{:,.0f}" for c, f in zip(df.columns, fm) if f}
            tab.dataframe(df.style.format(fmt), hide_index=True, use_container_width=True)


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
            tab.dataframe(df.style.format(fmt), hide_index=True, use_container_width=True)


def main():
    import streamlit as st
    st.set_page_config(page_title="D+0 Analytics", layout="wide")
    if not login(st):
        return
    pages = {"Region wise D+0 Reports": page_generate, "HQ D+0 Reports": page_hq,
             "Generate NDC / DC Reports": page_ndc_dc,
             "DCs Redirection Comparison": page_day, "Compare Two Periods": page_compare}
    pages[st.sidebar.radio("Select", list(pages))](st)


if __name__ == "__main__":
    main()

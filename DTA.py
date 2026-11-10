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
    page = st.sidebar.radio("Select", ["Generate D+0 Reports", "Compare Two Periods"])
    (page_generate if page.startswith("Generate") else page_compare)(st)


if __name__ == "__main__":
    main()

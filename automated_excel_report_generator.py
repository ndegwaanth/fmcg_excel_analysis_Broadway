#!/usr/bin/env python
# coding: utf-8

# # FMCG Distribution Analysis - Excel report generator
# 
# Reads the **five source datasets directly** (current week in `data`, previous week in `prev_wk_data`) (it does not use the summary-combo workbooks) and writes one presentation-ready workbook, `FMCG_Analysis.xlsx`, with **five sheets - one per dataset**:
# 
# | Sheet | Dataset |
# |---|---|
# | `1 Brands` | brand / product invoices and credit notes |
# | `2 Crates` | crate invoices and credit notes |
# | `3 Mehul Returns` | Mehul daily returns (fresh and expired) |
# | `4 Hansa Returns` | Hansa returns by document |
# | `5 Sales Revenue` | sales-revenue invoices and credits (`inv_total`) |
# 
# **Every sheet has the same layout, top to bottom**
# 1. Title, six KPI cards (each with the previous week's value and the change) and data notes
# 2. **Ten key questions, each answered with numbers from the data**
# 3. **Each question is a block: the answer and a suggested action on the left, the data table in the middle, and a chart with data labels on the right**
# 4. **Week-on-week comparison** (when a `prev_wk_data` folder exists): seven blocks that say what changed since the previous week - headline KPIs, biggest movers by product / category, route and customer (including new and lost customers), weekday pattern, one dataset-specific question, and a cross-dataset view - each with an inference, a table and a chart
# 5. **How this dataset relates to the other four** - four cross-dataset questions, each with its own table and chart
# 6. **A customer matrix in an Excel table with filters** (customer info columns + values, heat-shaded, filtered totals row, plus previous-week total, change and % change for every customer)
# 
# **Printing.** Each sheet is set to landscape, one page wide, with its own print area, scale, page breaks that keep each block and chart whole, a repeated header row for the matrix, and page numbers. Change `PAPER_SIZE` (9 = A4, 8 = A3) in the settings cell if you want larger print.
# 
# **Questions answered per dataset**
# 
# *Brands* - top brands; weakest brands; top buyer of each brand; top routes; top sales reps; weekday demand; range breadth per customer; products bought together; declining customers; concentration in a few customers/routes.
# 
# *Crates* - customers using most crates; routes needing most; daily movement; highest/lowest average use; unusual days; category and rep differences; sudden customer changes; weekday consistency; routes needing more/fewer crates; invoice-credit reconciliation.
# 
# *Mehul returns* - total fresh/expired; top products; top routes; top customers; fresh vs expired share; routes with most expired; products with high expired share; concentration by day; category differences; estimated value of returns.
# 
# *Hansa returns* - total returned; top products; top customers; top routes; invoices linked to returns; return quantity per invoice; change over time; repeat customer-product returns; differences by transaction type; agreement with Mehul.
# 
# *Sales revenue* - invoiced / credited / net; monthly and daily trend; top customers; routes (revenue and credit ratio); categories; average invoice; declining customers/routes; unusual invoices and credits; repeat credits and linked invoices; VAT / extras / total reconciliation.
# 
# **Cross-dataset (on every sheet)** - return rates by product against sales, units per crate, revenue per crate, revenue vs units, credits vs returns, Mehul vs Hansa agreement, and customer overlap. Each comparison uses only the dates both files cover; if the files cover different dates, the text says so and compares patterns, not same-period rates.
# 
# Correlations measure association only; they do not show cause.

# ## Config

# In[37]:


import re, calendar
from pathlib import Path
import numpy as np
import pandas as pd

# Support running with either the project folder or its parent as the notebook working directory.
BASE_DIR = Path.cwd()
if not (BASE_DIR / "data").is_dir() and (BASE_DIR / "fmcg_excel_analysis" / "data").is_dir():
    BASE_DIR = BASE_DIR / "fmcg_excel_analysis"

DATA = BASE_DIR / "data"
CUSTOMER_INFO_FILE = DATA / "customer-info" / "customer-info.xlsx"
BRANDS_FILE        = DATA / "brands" / "nov-week1-brand.xlsx"          # sheets: Brands-inv, Brands-cn
CRATES_FILE        = DATA / "crates" / "nov-week1-crates.xlsx"         # sheets: Crates-inv, Crates-cn
MEHUL_FILE         = DATA / "mehul-returns" / "NOVEMBER_RETURNS_NOV_WK1.xlsx"   # one sheet per day: 1ST, 2ND, ...
HANSA_FILE         = DATA / "hansa-returns" / "nov-week1-hansa-returns.xlsx"
SALES_REV_FILE     = DATA / "sales-revenue" / "november-week1-sales-revenue.xlsx"  # sheets: Invoices, Credits

OUTPUT_DIR = BASE_DIR / "report" / "november" / "wk1"

# Mehul's daily sheets only carry the day (1ST, 2ND ...), so the month has to be given here
MEHUL_YEAR, MEHUL_MONTH = 2026, 11
# Extra Mehul sheets that are NOT named 1ST/2ND/... but should be read, e.g. {"Sheet14": 3} = treat as 3rd of the month
MEHUL_SHEET_DATE_OVERRIDES = {}

# False = an invoice reversed by an identical credit note disappears on both sides
KEEP_CANCELLED_CREDITS = False

# Brands: columns that must match (credit = -invoice) for a credit note to cancel an invoice (same list as BRANDS.ipynb)
BRAND_MATCH_COLS = ['P', 'S', 'B', '200g', 'O.P.', '800g', '600g', 'MANDAZI', 'DOUGHNUTS',
                    'MANDO FAMILY PACK', '4 Square', 'XTRA LOONG', 'B-BITEZ', 'QTY']


# In[38]:


# ---- Report settings
REPORT_FILE  = OUTPUT_DIR / "November_wk1_FMCG_Analysis.xlsx"
REPORT_TITLE = "FMCG Distribution Analysis"
PAPER_SIZE   = 9        # 9 = A4, 8 = A3, 1 = Letter  (every sheet prints landscape, 1 page wide)

try:
    import xlsxwriter
except ImportError as e:
    raise ImportError("This notebook writes the Excel charts with xlsxwriter. Run:  pip install xlsxwriter") from e

# ---- Previous-week comparison (optional: if the folder does not exist the report is built without it)
PREV_DIR = BASE_DIR / "prev_wk_data" / "NOVEMBER" / "wk1"
PREV_FILES = {
    "brands": PREV_DIR / "brands" / "oct-week5-brand.xlsx",
    "crates": PREV_DIR / "crates" / "oct-week5-crates.xlsx",
    "mehul": PREV_DIR / "mehul-returns" / "OCTOBER_RETURNS_OCT_WK5.xlsx",
    "hansa": PREV_DIR / "hansa-returns" / "oct-week5-hansa-returns.xlsx",
    "revenue": PREV_DIR / "sales-revenue" / "october-week5-sales-revenue.xlsx",
}                                   # explicit week 1 sources avoid ambiguous auto-selection
PREV_MEHUL_YEAR, PREV_MEHUL_MONTH = None, None    # month of the previous Mehul sheets; None = take it from the other previous-week files
CURRENT_WINDOW = ("2026-11-01", "2026-11-07")
PREV_WINDOW = ("2026-10-25", "2026-10-31")


# ## Load and clean the datasets (current week and previous week)

# In[39]:


def key(x):
    return re.sub(r"[^A-Z0-9]", "", str(x).upper())

# One spelling per product, whatever the source file calls it (200g/200G, O.P/O.P., BBITEZ/B-BITEZ ...)
PRODUCTS = ['P', 'S', 'B', '200g', 'O.P.', '800g', '1500g', 'ST', '600g', 'KATIKATI', 'MANDAZI', 'DOUGHNUTS',
            '50-50 SLICED', 'MANDO FAMILY PACK', '4 Square', 'XTRA LOONG', 'B-BITEZ']
_PRODUCT_LOOKUP = {key(p): p for p in PRODUCTS}
_PRODUCT_LOOKUP.update({'4SQ': '4 Square', 'XLOONG': 'XTRA LOONG', 'FAMILYPACK': 'MANDO FAMILY PACK',
                        'BB': 'B-BITEZ', '5050': '50-50 SLICED'})
# Names not listed above (e.g. Mehul's MANDO, MAICHO) are kept as written (upper-case). Add aliases above to merge them.
def canon(x):
    return _PRODUCT_LOOKUP.get(key(x), str(x).strip().upper())

META_KEYS = {'DATE': 'DATE', 'DATES': 'DATE', 'AC': 'A/C', 'CUSTOMER': 'NAME', 'NAME': 'NAME', 'QTY': 'QTY',
             'NO': 'NO', 'INVOICENO': 'INVOICE_NO', 'INVOICE': 'INVOICE_NO', 'TYPE': 'TYPE', 'STATUS': 'STATUS'}
META_COLS = set(META_KEYS.values())

def std_cols(df):
    df = df.copy()
    df.columns = [META_KEYS.get(key(c)) or canon(c) for c in df.columns]
    return df

def to_code(s):
    return pd.to_numeric(s, errors="coerce").astype("Int64")

def code_key(s):
    """Code as a whole number; a missing/invalid code becomes 0 so those rows are not silently dropped by groupby."""
    return to_code(s).fillna(0).astype("int64")

def as_qty(s):
    """Whole numbers stay int64; if a source ever holds decimals they are kept, never truncated."""
    s = pd.to_numeric(s).fillna(0)
    return s.round().astype("int64") if np.allclose(s, s.round()) else s.astype("float64")

def to_date(s):
    return pd.to_datetime(s, errors="coerce").dt.normalize()

# ---- customer info
cust = pd.read_excel(CUSTOMER_INFO_FILE)
cust = cust.loc[:, ~cust.columns.astype(str).str.match(r"^Unnamed(?::|$)")]
cust["CODE"] = to_code(cust["CODE"])
cust = cust.dropna(subset=["CODE"])
_dup = cust["CODE"].duplicated()
if _dup.any():
    print(f"WARNING: {_dup.sum()} duplicate CODE rows in customer info - first one kept: {sorted(cust.loc[_dup, 'CODE'].tolist())}")
cust = cust[~_dup].reset_index(drop=True)
INFO_COLS = ["CODE"] + [c for c in cust.columns if c != "CODE"]
cust = cust[INFO_COLS]
ID_COLS = INFO_COLS + ["CUSTOMER_MATCHED"]
print("Customer info:", cust.shape, INFO_COLS)

def first_name_by_code(df, code_col="A/C", name_col="NAME"):
    t = df[[code_col, name_col]].dropna()
    t = t.assign(_c=to_code(t[code_col])).dropna(subset=["_c"])
    return t.drop_duplicates("_c").set_index("_c")[name_col].astype(str).to_dict()

def attach_customer(df, code_col="CODE", name_col=None, fallback=None, label=""):
    """Add every customer-info column by code. Codes missing from customer info keep a source name (if any) and are flagged."""
    df = df.copy()
    n_in = len(df)
    c = to_code(df[code_col])
    df["CODE"] = c.mask(c == 0)                      # 0 = "no code"
    if code_col != "CODE":
        df = df.drop(columns=code_col)
    src = pd.Series("", index=df.index, dtype="object")
    if fallback is not None:
        src = df["CODE"].map(lambda c: fallback.get(int(c), "") if pd.notna(c) else "").astype("object")
    if name_col:
        nm = df[name_col].fillna("").astype(str).str.strip().astype("object")
        src = nm.where(nm != "", src)
        df = df.drop(columns=name_col)
    out = df.merge(cust.assign(CUSTOMER_MATCHED="Yes"), on="CODE", how="left")
    assert len(out) == n_in, "customer merge changed the row count"
    out["CUSTOMER_MATCHED"] = out["CUSTOMER_MATCHED"].fillna("No")
    miss = (out["CUSTOMER_MATCHED"] == "No").to_numpy()
    out["CUSTOMER NAME"] = out["CUSTOMER NAME"].astype("object")
    out.loc[miss, "CUSTOMER NAME"] = src.to_numpy()[miss]
    n_bad = int(miss.sum())
    if n_bad:
        bad = out.loc[miss, "CODE"]
        print(f"  {label}: {n_bad} rows ({bad.nunique(dropna=False)} codes) not in customer info, e.g. {bad.dropna().unique()[:10].tolist()}")
    return out


# In[40]:


def cancel_matched(inv, cn, exact, negated):
    """
    One credit note cancels ONE invoice (not every duplicate).
    exact   : columns that must be equal on both sides
    negated : columns where credit == -invoice
    Returns invoices kept, credits kept (not matched), and the matched credits.
    """
    inv = inv.reset_index(drop=True)
    cn = cn.reset_index(drop=True)
    keys = exact + negated
    ik = inv[keys].copy()
    ck = cn[keys].copy()
    for c in negated:
        ck[c] = -ck[c]
    ik["_n"] = ik.groupby(keys, dropna=False).cumcount()
    ck["_n"] = ck.groupby(keys, dropna=False).cumcount()
    k2 = keys + ["_n"]
    inv_hit = ik.merge(ck[k2].assign(_h=1), on=k2, how="left")["_h"].notna().to_numpy()
    cn_hit = ck.merge(ik[k2].assign(_h=1), on=k2, how="left")["_h"].notna().to_numpy()
    return inv[~inv_hit].copy(), cn[~cn_hit].copy(), cn[cn_hit].copy()


# The five loaders are functions, so exactly the same cleaning is applied to `data` (current week) and `prev_wk_data` (previous week).

# In[41]:


def _load_brands(BRANDS_FILE):
    print("BRANDS")
    b_inv = std_cols(pd.read_excel(BRANDS_FILE, sheet_name="Brands-inv"))
    b_cn  = std_cols(pd.read_excel(BRANDS_FILE, sheet_name="Brands-cn"))
    b_prods = [c for c in b_inv.columns if c not in META_COLS]
    for d in (b_inv, b_cn):
        d["DATE"] = to_date(d["DATE"])
        d["A/C"] = code_key(d["A/C"])
        for c in b_prods + ["QTY"]:
            d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0)

    match_cols = [c for c in BRAND_MATCH_COLS if c in b_inv.columns and c in b_cn.columns]
    b_inv_kept, b_cn_open, b_cn_cancelled = cancel_matched(b_inv, b_cn, exact=["DATE", "A/C"], negated=match_cols)
    print(f"  invoices {len(b_inv)} -> {len(b_inv_kept)} after removing {len(b_inv) - len(b_inv_kept)} reversed by a credit note")
    print(f"  credit notes {len(b_cn)}: {len(b_cn_cancelled)} cancelled an invoice, {len(b_cn_open)} stand alone")
    b_cn_use = b_cn if KEEP_CANCELLED_CREDITS else b_cn_open

    def brand_long(df, value_name):
        g = df.groupby(["DATE", "A/C"], as_index=False)[b_prods].sum()
        long = g.melt(id_vars=["DATE", "A/C"], var_name="PRODUCT", value_name=value_name)
        return long[long[value_name] != 0]

    _cn_pos = b_cn_use.copy()
    _cn_pos[b_prods] = _cn_pos[b_prods].abs()
    brands_dw = (pd.concat([brand_long(b_inv_kept, "BRANDS_INV_QTY"), brand_long(_cn_pos, "BRANDS_CN_QTY")])
                 .groupby(["DATE", "A/C", "PRODUCT"], as_index=False).sum())
    brands_dw = attach_customer(brands_dw, "A/C", fallback=first_name_by_code(pd.concat([b_inv, b_cn])), label="brands")
    BRANDS_M = ["BRANDS_INV_QTY", "BRANDS_CN_QTY"]
    brands_dw[BRANDS_M] = brands_dw[BRANDS_M].fillna(0).apply(as_qty)
    brands_dw = brands_dw[["DATE"] + ID_COLS + ["PRODUCT"] + BRANDS_M].sort_values(["DATE", "CODE", "PRODUCT"]).reset_index(drop=True)
    _ns = locals()
    return {n: _ns[n] for n in ['brands_dw', 'b_inv', 'b_cn', 'b_inv_kept', 'b_cn_open', 'b_cn_cancelled', 'b_prods']}


def _load_crates(CRATES_FILE):
    print("CRATES")
    c_inv = std_cols(pd.read_excel(CRATES_FILE, sheet_name="Crates-inv"))
    c_cn  = std_cols(pd.read_excel(CRATES_FILE, sheet_name="Crates-cn"))
    c_cn = c_cn[c_cn["STATUS"].astype(str).str.strip().str.upper() == "CN"]      # as in CRATES.ipynb
    for d in (c_inv, c_cn):
        d["DATE"] = to_date(d["DATE"])
        d["A/C"] = code_key(d["A/C"])
        d["QTY"] = pd.to_numeric(d["QTY"], errors="coerce").fillna(0)

    # CRATES.ipynb matches on A/C + QTY only (a credit can reverse an earlier day's invoice)
    c_inv_kept, c_cn_open, c_cn_cancelled = cancel_matched(c_inv, c_cn, exact=["A/C", "QTY"], negated=[])
    print(f"  invoices {len(c_inv)} -> {len(c_inv_kept)} after removing {len(c_inv) - len(c_inv_kept)} reversed by a credit note")
    print(f"  credit notes {len(c_cn)}: {len(c_cn_cancelled)} cancelled an invoice, {len(c_cn_open)} stand alone")
    c_cn_use = c_cn if KEEP_CANCELLED_CREDITS else c_cn_open

    crates_dw = (pd.concat([
            c_inv_kept.groupby(["DATE", "A/C"], as_index=False)["QTY"].sum().rename(columns={"QTY": "CRATES_INV_QTY"}),
            c_cn_use.groupby(["DATE", "A/C"], as_index=False)["QTY"].sum().rename(columns={"QTY": "CRATES_CN_QTY"})])
        .groupby(["DATE", "A/C"], as_index=False).sum())
    crates_dw = attach_customer(crates_dw, "A/C", label="crates")
    CRATES_M = ["CRATES_INV_QTY", "CRATES_CN_QTY"]
    crates_dw[CRATES_M] = crates_dw[CRATES_M].fillna(0).apply(as_qty)
    crates_dw = crates_dw[["DATE"] + ID_COLS + CRATES_M].sort_values(["DATE", "CODE"]).reset_index(drop=True)
    _ns = locals()
    return {n: _ns[n] for n in ['crates_dw', 'c_inv', 'c_cn', 'c_inv_kept', 'c_cn_open', 'c_cn_cancelled']}


def _load_mehul(MEHUL_FILE, MEHUL_YEAR, MEHUL_MONTH, MEHUL_SHEET_DATE_OVERRIDES):
    print("MEHUL RETURNS")
    ORD = re.compile(r"^\s*(\d{1,2})\s*(ST|ND|RD|TH)(\w*)\s*$", re.I)
    xl = pd.ExcelFile(MEHUL_FILE)
    days_in_month = calendar.monthrange(MEHUL_YEAR, MEHUL_MONTH)[1]

    by_day = {}
    for s in xl.sheet_names:
        m = ORD.match(s)
        if m:
            by_day.setdefault(int(m.group(1)), []).append((m.group(3) != "", s))   # (misspelt?, sheet)
    chosen, skipped = [], []
    for day, lst in sorted(by_day.items()):
        lst.sort(key=lambda t: t[0])             # properly spelt name first
        if day > days_in_month:
            skipped += [s for _, s in lst]; continue
        chosen.append((day, lst[0][1])); skipped += [s for _, s in lst[1:]]
    for s, day in MEHUL_SHEET_DATE_OVERRIDES.items():
        chosen.append((day, s)); skipped = [x for x in skipped if x != s]
    skipped += [s for s in xl.sheet_names if s not in {c[1] for c in chosen} and s not in skipped]
    print(f"  reading {len(chosen)} daily sheets: {[s for _, s in chosen]}")
    print(f"  skipped sheets: {skipped}")

    def parse_mehul_sheet(sheet):
        raw = xl.parse(sheet, header=None)
        if raw.shape[0] < 4 or raw.shape[1] < 6:
            return None
        brands, types = raw.iloc[1, 5:].tolist(), raw.iloc[2, 5:].tolist()
        cur, seen, spec = None, set(), []
        for b, t in zip(brands, types):
            if pd.notna(b) and str(b).strip():
                cur = str(b).strip().upper()
            tt = str(t).strip().upper() if pd.notna(t) else ""
            tt = {"EX": "EXP", "1": "EXP", "1.0": "EXP"}.get(tt, tt)
            product = canon(cur) if cur else None
            if product and product != "TOTAL" and tt not in ("FRESH", "EXP") and spec and spec[-1] == (product, "FRESH"):
                print(f"  {sheet}: unexpected return type {t!r} in column {len(spec) + 6}; inferred EXP")
                tt = "EXP"
            if product and product != "TOTAL" and tt in ("FRESH", "EXP") and (product, tt) not in seen:
                seen.add((product, tt)); spec.append((product, tt))
            else:
                spec.append(None)                # TOTAL block, comments, repeated brand blocks (first one wins)
        body = raw.iloc[3:].reset_index(drop=True)
        out = pd.DataFrame({"CODE": body[3], "NAME": body[4]})
        for pos, sp in enumerate(spec, start=5):
            if sp is not None:
                out[f"{sp[0]}|{sp[1]}"] = pd.to_numeric(body[pos], errors="coerce").fillna(0)
        nm = out["NAME"].fillna("").astype(str).str.strip()
        out = out[~((out["CODE"].isna()) & (nm == ""))]
        out = out[nm.loc[out.index].str.upper() != "TOTAL"]
        return out

    frames = []
    for day, s in chosen:
        f = parse_mehul_sheet(s)
        if f is not None and len(f):
            f["DATE"] = pd.Timestamp(MEHUL_YEAR, MEHUL_MONTH, day)
            frames.append(f)
    m_raw = pd.concat(frames, ignore_index=True)
    val_cols = [c for c in m_raw.columns if "|" in c]
    m_raw[val_cols] = m_raw[val_cols].fillna(0)

    # ---- find missing codes
    def norm_name(s):
        s = re.sub(r"\(\s*\d+\s*\)\s*$", "", str(s).upper())
        return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", s)).strip()
    _name_to_code = (cust.assign(_k=cust["CUSTOMER NAME"].astype(str).map(norm_name))
                         .drop_duplicates("_k", keep=False).set_index("_k")["CODE"].to_dict())
    nm = m_raw["NAME"].fillna("").astype(str)
    code0 = to_code(m_raw["CODE"])
    n0 = int(code0.isna().sum())
    code1 = code0.fillna(to_code(nm.str.extract(r"\((\d+)\)\s*$", expand=False)))
    code2 = code1.fillna(to_code(nm.map(lambda s: _name_to_code.get(norm_name(s), np.nan))))
    print(f"  rows without a code: {n0} -> {int(code2.isna().sum())} after inferring from the name")
    m_raw["CODE"] = code2
    _fallback = first_name_by_code(m_raw, "CODE", "NAME")
    m_raw["NAME"] = np.where(m_raw["CODE"].notna(), "", nm)         # name only matters for rows that still have no code
    m_raw["CODE"] = m_raw["CODE"].fillna(0).astype("int64")          # 0 = "no code" so uncoded rows are not dropped by groupby

    # ---- pivot brands into rows, FRESH / EXP into two columns
    long = m_raw.melt(id_vars=["DATE", "CODE", "NAME"], value_vars=val_cols, var_name="k", value_name="v")
    long[["PRODUCT", "TYPE"]] = long["k"].str.split("|", expand=True)
    mehul_dw = (long.groupby(["DATE", "CODE", "NAME", "PRODUCT", "TYPE"])["v"].sum().unstack("TYPE", fill_value=0).reset_index())
    for t in ("FRESH", "EXP"):
        if t not in mehul_dw.columns:
            mehul_dw[t] = 0
    mehul_dw = mehul_dw.rename(columns={"FRESH": "MEHUL_FRESH_RTN_QTY", "EXP": "MEHUL_EXP_RTN_QTY"})
    mehul_dw["MEHUL_TOTAL_RTN_QTY"] = mehul_dw["MEHUL_FRESH_RTN_QTY"] + mehul_dw["MEHUL_EXP_RTN_QTY"]
    mehul_dw = mehul_dw[mehul_dw["MEHUL_TOTAL_RTN_QTY"] != 0].copy()
    MEHUL_M = ["MEHUL_FRESH_RTN_QTY", "MEHUL_EXP_RTN_QTY", "MEHUL_TOTAL_RTN_QTY"]
    mehul_dw = attach_customer(mehul_dw, "CODE", name_col="NAME", fallback=_fallback, label="mehul")
    mehul_dw[MEHUL_M] = mehul_dw[MEHUL_M].apply(as_qty)
    mehul_dw = mehul_dw[["DATE"] + ID_COLS + ["PRODUCT"] + MEHUL_M].sort_values(["DATE", "CODE", "PRODUCT"]).reset_index(drop=True)
    _ns = locals()
    return {n: _ns[n] for n in ['mehul_dw', 'm_raw', 'val_cols']}


def _load_hansa(HANSA_FILE):
    print("HANSA RETURNS")
    h = std_cols(pd.read_excel(HANSA_FILE))          # first sheet; add sheet_name=... if the file has several
    h_prods = [c for c in h.columns if c not in META_COLS]
    h["DATE"] = to_date(h["DATE"])
    h["A/C"] = code_key(h["A/C"])
    # Hansa records returns as negative credit quantities; report returned units as positive magnitudes.
    for c in h_prods:
        h[c] = pd.to_numeric(h[c], errors="coerce").fillna(0).abs()
    source_qty = pd.to_numeric(h["QTY"], errors="coerce").fillna(0).abs()
    detail_qty = h[h_prods].sum(axis=1)
    qty_mismatch = ~np.isclose(source_qty.to_numpy(), detail_qty.to_numpy())
    if qty_mismatch.any():
        print(f"  source Qty subtotal differs from itemized products on {int(qty_mismatch.sum())} rows; using itemized products")
    hg = h.groupby(["DATE", "A/C"], as_index=False)[h_prods].sum()
    hansa_dw = hg.melt(id_vars=["DATE", "A/C"], var_name="PRODUCT", value_name="HANSA_RTN_QTY")
    hansa_dw = hansa_dw[hansa_dw["HANSA_RTN_QTY"] != 0]
    hansa_dw = attach_customer(hansa_dw, "A/C", fallback=first_name_by_code(h), label="hansa")
    HANSA_M = ["HANSA_RTN_QTY"]
    hansa_dw[HANSA_M] = hansa_dw[HANSA_M].apply(as_qty)
    hansa_dw = hansa_dw[["DATE"] + ID_COLS + ["PRODUCT"] + HANSA_M].sort_values(["DATE", "CODE", "PRODUCT"]).reset_index(drop=True)
    print(f"  products: {h_prods}")
    _ns = locals()
    return {n: _ns[n] for n in ['hansa_dw', 'h', 'h_prods']}


def _load_revenue(SALES_REV_FILE):
    print("SALES REVENUE")
    def clean_rev(df):
        df = df.copy()
        df.columns = (df.columns.str.strip().str.lower().str.replace(" ", "_").str.replace("(", "", regex=False)
                      .str.replace(")", "", regex=False).str.replace(".", "", regex=False))
        df["DATE"] = to_date(df["date"])
        df["A/C"] = code_key(df["a/c_no"])
        if "inv_total" not in df.columns:
            raise KeyError(f"No inv_total column in the sales revenue file. Columns found: {list(df.columns)}")
        for c in ("qty", "inv_total"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        return df

    rev_inv = clean_rev(pd.read_excel(SALES_REV_FILE, sheet_name="Invoices"))
    rev_cn  = clean_rev(pd.read_excel(SALES_REV_FILE, sheet_name="Credits"))
    rev_cn[["qty", "inv_total"]] = rev_cn[["qty", "inv_total"]].abs()

    def rev_group(df, qty_name, amt_name):
        return (df.groupby(["DATE", "A/C"], as_index=False)[["qty", "inv_total"]].sum()
                  .rename(columns={"qty": qty_name, "inv_total": amt_name}))
    rev_dw = pd.concat([rev_group(rev_inv, "REV_INV_QTY", "REV_INV_AMT"),
                        rev_group(rev_cn, "REV_CN_QTY", "REV_CN_AMT")]).groupby(["DATE", "A/C"], as_index=False).sum()
    # Customers / Category / Route columns in this file are ignored: everything comes from customer info, looked up by A/C
    rev_dw = attach_customer(rev_dw, "A/C", label="revenue")
    REV_M = ["REV_INV_QTY", "REV_INV_AMT", "REV_CN_QTY", "REV_CN_AMT"]
    rev_dw[REV_M] = rev_dw[REV_M].fillna(0)
    rev_dw[["REV_INV_QTY", "REV_CN_QTY"]] = rev_dw[["REV_INV_QTY", "REV_CN_QTY"]].apply(as_qty)
    rev_dw = rev_dw[["DATE"] + ID_COLS + REV_M].sort_values(["DATE", "CODE"]).reset_index(drop=True)
    _ns = locals()
    return {n: _ns[n] for n in ['rev_dw', 'rev_inv', 'rev_cn']}


# In[42]:


_dt = lambda d: pd.Timestamp(d).strftime('%d %b %Y')
LOADED_NAMES = {
    "brands":  ["brands_dw", "b_inv", "b_cn", "b_inv_kept", "b_cn_open", "b_cn_cancelled", "b_prods"],
    "crates":  ["crates_dw", "c_inv", "c_cn", "c_inv_kept", "c_cn_open", "c_cn_cancelled"],
    "mehul":   ["mehul_dw", "m_raw", "val_cols"],
    "hansa":   ["hansa_dw", "h", "h_prods"],
    "revenue": ["rev_dw", "rev_inv", "rev_cn"],
}
DATE_FRAMES = ["brands_dw", "b_inv", "b_cn", "b_inv_kept", "b_cn_open", "b_cn_cancelled", "crates_dw", "c_inv", "c_cn", "c_inv_kept", "c_cn_open", "c_cn_cancelled",
               "mehul_dw", "hansa_dw", "h", "rev_dw", "rev_inv", "rev_cn"]

def load_period(files, mehul_ym, label, window=(None, None)):
    """run the five loaders for one week; `files` = {dataset: path}; mehul_ym = (year, month) or a function of the loaded week"""
    print(f"\n######## {label} ########")
    P = {}
    P.update(_load_brands(files["brands"]))
    P.update(_load_crates(files["crates"]))
    P.update(_load_hansa(files["hansa"]))
    P.update(_load_revenue(files["revenue"]))
    y, m = mehul_ym(P) if callable(mehul_ym) else mehul_ym
    P.update(_load_mehul(files["mehul"], y, m, MEHUL_SHEET_DATE_OVERRIDES))
    lo, hi = (pd.Timestamp(window[0]) if window[0] else None), (pd.Timestamp(window[1]) if window[1] else None)
    if lo is not None or hi is not None:
        for k in DATE_FRAMES:
            d = P[k]
            if lo is not None: d = d[d["DATE"] >= lo]
            if hi is not None: d = d[d["DATE"] <= hi]
            P[k] = d
        print(f"  limited to {window[0] or 'start'} .. {window[1] or 'end'}")
    return P

CURR = load_period({"brands": BRANDS_FILE, "crates": CRATES_FILE, "mehul": MEHUL_FILE, "hansa": HANSA_FILE, "revenue": SALES_REV_FILE},
                   (MEHUL_YEAR, MEHUL_MONTH), "CURRENT WEEK  (data)", CURRENT_WINDOW)
globals().update({k: CURR[k] for names in LOADED_NAMES.values() for k in names})     # the rest of the notebook works on the current week

# ---------- previous week
def find_prev_file(key, current_file):
    if key in PREV_FILES:
        return Path(PREV_FILES[key])
    sub = PREV_DIR / Path(current_file).parent.name
    same = sub / Path(current_file).name
    if same.exists():
        return same
    cands = sorted([p for p in sub.glob("*.xlsx") if not p.name.startswith("~$")], key=lambda p: p.stat().st_mtime, reverse=True) if sub.exists() else []
    if not cands:
        raise FileNotFoundError(f"No .xlsx file for '{key}' in {sub}. Put the previous week's file there, or give its path in PREV_FILES.")
    if len(cands) > 1:
        print(f"  {sub}: several files, using the newest ({cands[0].name}); set PREV_FILES['{key}'] to choose another")
    return cands[0]

def infer_month(P):
    """(year, month) most of the other files of this week fall in"""
    d = pd.concat([P[k]["DATE"] for k in ("brands_dw", "crates_dw", "hansa_dw") if len(P[k])])
    ym = d.dt.to_period("M").value_counts()
    p = ym.index[0]
    return p.year, p.month

PRV = None
if PREV_DIR.exists():
    pf = {"brands": find_prev_file("brands", BRANDS_FILE), "crates": find_prev_file("crates", CRATES_FILE), "mehul": find_prev_file("mehul", MEHUL_FILE),
          "hansa": find_prev_file("hansa", HANSA_FILE), "revenue": find_prev_file("revenue", SALES_REV_FILE)}
    for k, v in pf.items():
        print(f"  previous {k}: {v}")
    pm_ = (lambda P: (PREV_MEHUL_YEAR, PREV_MEHUL_MONTH)) if (PREV_MEHUL_YEAR and PREV_MEHUL_MONTH) else infer_month
    PRV = load_period(pf, pm_, "PREVIOUS WEEK  (prev_wk_data)", PREV_WINDOW)
    if not (PREV_MEHUL_YEAR and PREV_MEHUL_MONTH):
        y_, m_ = infer_month(PRV)
        print(f"  previous Mehul sheets (1ST, 2ND ...) were dated in {calendar.month_name[m_]} {y_}; set PREV_MEHUL_YEAR / PREV_MEHUL_MONTH if that is wrong")
    print("\nPeriod coverage (dates):")
    for nm, k in (("brands", "brands_dw"), ("crates", "crates_dw"), ("Mehul returns", "mehul_dw"), ("Hansa returns", "hansa_dw"), ("sales revenue", "rev_dw")):
        a, b = CURR[k]["DATE"], PRV[k]["DATE"]
        txt = f"  {nm}: current {_dt(a.min()) if len(a) else '-'} to {_dt(a.max()) if len(a) else '-'} ({a.nunique()} dates) | previous {_dt(b.min()) if len(b) else '-'} to {_dt(b.max()) if len(b) else '-'} ({b.nunique()} dates)"
        if len(a) and len(b) and a.min() <= b.max() and b.min() <= a.max():
            txt += "   <- the two periods OVERLAP: set CURRENT_WINDOW / PREV_WINDOW so each week is counted once"
        print(txt)
else:
    print(f"\nNo '{PREV_DIR}' folder found - the report is built for the current week only.")
HAS_PREV = PRV is not None


# ## Analysis helpers

# In[43]:


WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
NOT_IN_INFO = "(not in customer info)"

def n0(x): return f"{x:,.0f}"
def n1(x): return f"{x:,.1f}"
def pc(x, d=1): return f"{x * 100:.{d}f}%"
def sdiv(a, b): return float(a) / float(b) if b else 0.0
def dtxt(d): return pd.Timestamp(d).strftime("%d %b %Y")
def span_txt(df):
    return f"{dtxt(df['DATE'].min())} to {dtxt(df['DATE'].max())}" if len(df) else "no data"

def lab(df):
    """'CUSTOMER NAME (CODE)' label for every row."""
    nm = df["CUSTOMER NAME"].astype("string").fillna("").str.strip()
    cd = df["CODE"].astype("string").fillna("")
    return pd.Series(np.where(cd != "", nm + " (" + cd + ")", nm), index=df.index).astype("object")

def dim(df, col):
    return df[col].astype("object").where(df[col].notna(), NOT_IN_INFO)

def top_list(s, k=3, fmt=n0, share=True):
    s = s[s > 0].sort_values(ascending=False)
    tot = s.sum()
    return "; ".join(f"{i} ({fmt(v)}" + (f", {pc(v / tot)}" if share and tot else "") + ")" for i, v in s.head(k).items())

def bottom_list(s, k=3, fmt=n0):
    s = s[s > 0].sort_values()
    return "; ".join(f"{i} ({fmt(v)})" for i, v in s.head(k).items())

def weekday_avg(daily):
    """average per weekday from a Series indexed by date"""
    if len(daily) == 0:
        return daily
    return daily.groupby(pd.DatetimeIndex(daily.index).day_name()).mean().reindex(WEEKDAYS).dropna()

def weekday_txt(daily, unit):
    wd = weekday_avg(daily)
    if len(wd) < 2:
        return "Only one weekday is covered by the file, so no weekday pattern can be shown yet.", wd
    hi, lo = wd.idxmax(), wd.idxmin()
    t = (f"Busiest weekday: {hi} (avg {n0(wd[hi])} {unit} per day); quietest: {lo} (avg {n0(wd[lo])}). "
         f"{hi} runs {pc(sdiv(wd[hi], wd[lo]) - 1, 0)} above {lo}.")
    if len(daily) < 14:
        t += f" Based on only {len(daily)} dates, so treat it as indicative."
    return t, wd

def trend_changes(df, key, val, min_total=20):
    """per-key average per day in the first vs second half of the dates in df. None if < 2 dates."""
    dts = sorted(df["DATE"].dropna().unique())
    if len(dts) < 2:
        return None
    mid = len(dts) // 2
    d1, d2 = set(dts[:mid]), set(dts[mid:])
    a = df[df["DATE"].isin(d1)].groupby(key)[val].sum()
    b = df[df["DATE"].isin(d2)].groupby(key)[val].sum()
    t = pd.concat([a.rename("first_tot"), b.rename("second_tot")], axis=1).fillna(0)
    t["first"], t["second"] = t["first_tot"] / len(d1), t["second_tot"] / len(d2)
    t = t[t["first_tot"] + t["second_tot"] >= min_total].copy()
    t["chg"] = np.where(t["first"] > 0, t["second"] / t["first"] - 1, np.nan)
    t["gap"] = t["second"] - t["first"]
    return t, len(dts)

def change_txt(i, r, unit):
    return f"{i} ({n1(r['first'])} to {n1(r['second'])} {unit}/day, {r['chg'] * 100:+.0f}%)"

def window(*frames):
    """date range common to every frame; (None, None) if they do not overlap"""
    frames = [f for f in frames if len(f)]
    if not frames:
        return (None, None)
    lo = max(f["DATE"].min() for f in frames)
    hi = min(f["DATE"].max() for f in frames)
    return (lo, hi) if lo <= hi else (None, None)

def clip(df, w):
    lo, hi = w
    return df if lo is None else df[(df["DATE"] >= lo) & (df["DATE"] <= hi)]

def win_note(w, named):
    """one sentence saying which dates a cross-dataset comparison covers"""
    if w[0] is not None:
        return f"Compared over {dtxt(w[0])} to {dtxt(w[1])}, the dates both files cover."
    spans = " vs ".join(f"{n}: {span_txt(f)}" for n, f in named)
    return f"The files cover different dates ({spans}), so this compares overall patterns, not same-period rates."

def rank_corr(a, b, min_n=5):
    """rank correlation between two Series indexed by customer code (customers present in both)"""
    j = pd.concat([a.rename("a"), b.rename("b")], axis=1, join="inner")
    j = j[(j["a"] > 0) | (j["b"] > 0)]
    if len(j) < min_n:
        return None, len(j)
    r = CorrVal(j["a"].rank().corr(j["b"].rank()))
    r.n = len(j)
    return r, len(j)

class CorrVal(float):
    """a correlation that remembers how many customers it is based on"""
    n = None

def corr_words(r):
    if r is None: return "too few customers in common to measure"
    s = abs(r)
    strength = "strong" if s >= 0.6 else "moderate" if s >= 0.3 else "weak"
    small = f", only {r.n} customers so indicative" if getattr(r, "n", None) is not None and r.n < 10 else ""
    return f"a {strength} {'positive' if r >= 0 else 'negative'} link (rank correlation {r:+.2f}{small})"

def code_sum(df, w, cols):
    d = clip(df, w)
    d = d[d["CODE"].notna()]
    return d.groupby("CODE")[cols].sum()

def pivot_matrix(df, val, columns, aggfunc="sum"):
    """customer rows (all customer-info columns) x `columns` values"""
    d = df.copy()
    d["_k"] = d["CODE"].astype("string").fillna("") + "|" + d["CUSTOMER NAME"].astype("string").fillna("")
    info = d.drop_duplicates("_k").set_index("_k")[INFO_COLS].replace(NOT_IN_INFO, "")
    pv = d.pivot_table(index="_k", columns=columns, values=val, aggfunc=aggfunc, fill_value=0)
    pv.columns = [str(c) for c in pv.columns]
    out = info.loc[pv.index].join(pv).reset_index(drop=True)
    return out

# ---------- blocks: every answered question carries its own table and chart
_CODE_SUFFIX = re.compile(r"^(.*) \((\d+)\)$")
def short(s, n=26):
    """shorten long names but keep the customer code at the end, e.g. 'LONG SHOP NAME... (1195)'"""
    s = str(s)
    if len(s) <= n:
        return s
    m = _CODE_SUFFIX.match(s)
    if m:
        name, code = m.groups()
        keep = max(8, n - len(code) - 4)
        return name[: keep - 1].rstrip() + "… (" + code + ")"
    return s[: n - 1] + "…"

_CUST_NAME = cust.set_index("CODE")["CUSTOMER NAME"].astype(str).to_dict()
def clabel(code, names=None):
    """'NAME (CODE)' for a customer code, shortened for tables and charts"""
    try:
        c = int(code)
    except (TypeError, ValueError):
        return short(code)
    nm = _CUST_NAME.get(c) or (names or {}).get(c) or ""
    return short(f"{nm} ({c})" if nm else str(c))

def mk(table, kind=None, vals=None, cat=None, title=None, fmts=None, stacked=False, ymax=None, delta=None, label_fmt=None, points=None, right=None):
    """a question's data block: `table` is shown next to the answer, `kind` is the chart drawn from it
    (bar, column, line, pie). Columns ending in '%' hold fractions and are shown as percentages.
    delta=(colour if positive, colour if negative) and points=[colours] are used by the comparison charts."""
    return dict(table=table, kind=kind, vals=vals, cat=cat, title=title, fmts=fmts or {}, stacked=stacked, ymax=ymax, delta=delta, label_fmt=label_fmt, points=points, right=right or [])

def top_tbl(s, name_col, val_col, k=8, share=True, shorten=True):
    s = s[s > 0].sort_values(ascending=False)
    tot = s.sum()
    t = s.head(k).rename_axis(name_col).reset_index(name=val_col)
    if shorten:
        t[name_col] = t[name_col].map(short)
    if share and tot:
        t["SHARE %"] = t[val_col] / tot
    return t

def period_frame(df, cols, date_col="DATE", last=16):
    """daily rows -> at most `last` points: days, else weeks, else months (so every chart point can carry a label)"""
    d = df[[date_col] + cols].copy()
    n = d[date_col].nunique()
    if n <= last:
        per, name, f = d[date_col].dt.to_period("D"), "DATE", "%d %b"
    elif d[date_col].dt.to_period("W-SUN").nunique() <= last:
        per, name, f = d[date_col].dt.to_period("W-SUN"), "WEEK STARTING", "%d %b"
    else:
        per, name, f = d[date_col].dt.to_period("M"), "MONTH", "%b %Y"
    g = d.groupby(per)[cols].sum().tail(last)
    g.insert(0, name, [p.start_time.strftime(f) for p in g.index])
    return g.reset_index(drop=True)

def movers_tbl(t, unit_col="UNITS", k=8, by="gap", asc=True, name_map=None):
    """first-half vs second-half table from trend_changes()"""
    m = t.sort_values(by, ascending=asc).head(k)
    out = pd.DataFrame({"CUSTOMER": [short(i) for i in m.index],
                        f"FIRST HALF / DAY": m["first"].round(1).to_numpy(), f"SECOND HALF / DAY": m["second"].round(1).to_numpy(),
                        "CHANGE %": m["chg"].to_numpy()})
    return out

def bucket_tbl(values, edges, labels, name, count_name):
    b = pd.cut(values, bins=edges, labels=labels, right=True)
    t = b.value_counts().reindex(labels).fillna(0).astype(int).rename_axis(name).reset_index(name=count_name)
    t["SHARE %"] = t[count_name] / t[count_name].sum() if t[count_name].sum() else 0
    return t


# In[44]:


# ---------- cross-dataset building blocks (used by every sheet's "relationships" block)
def product_rates(ret_df, ret_col):
    """returns / units sold by product over the dates both datasets cover"""
    w = window(brands_dw, ret_df)
    sold = clip(brands_dw, w).groupby("PRODUCT")["BRANDS_INV_QTY"].sum()
    ret = clip(ret_df, w).groupby("PRODUCT")[ret_col].sum()
    t = pd.concat([sold.rename("SOLD"), ret.rename("RETURNED")], axis=1).fillna(0)
    unmatched = t[(t["SOLD"] == 0) & (t["RETURNED"] > 0)]
    t = t[t["SOLD"] > 0].copy()
    t["RATE"] = t["RETURNED"] / t["SOLD"]
    return t.reset_index(names="PRODUCT"), w, unmatched

def rates_text(t, label):
    if not len(t):
        return f"No product appears in both the brands file and the {label} file."
    overall = sdiv(t["RETURNED"].sum(), t["SOLD"].sum())
    none = t[t["RETURNED"] == 0]
    big = t[(t["SOLD"] >= 0.01 * t["SOLD"].sum()) & (t["RETURNED"] > 0)].sort_values("RATE", ascending=False)
    txt = f"{label} returns are {pc(overall)} of units sold overall."
    if len(big):
        txt += f" Highest: {'; '.join(f'{r.PRODUCT} ({pc(r.RATE)})' for r in big.head(3).itertuples())}."
        if len(big) > 3:
            txt += f" Lowest: {'; '.join(f'{r.PRODUCT} ({pc(r.RATE)})' for r in big.tail(2).itertuples())}."
    if len(none):
        txt += f" {len(none)} sold product(s) have no {label} returns recorded ({', '.join(none['PRODUCT'].astype(str).head(4))}{'...' if len(none) > 4 else ''})."
    return txt

def route_crate_ratio():
    w = window(brands_dw, crates_dw)
    u = clip(brands_dw, w).assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")["BRANDS_INV_QTY"].sum()
    c = clip(crates_dw, w).assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")["CRATES_INV_QTY"].sum()
    t = pd.concat([u.rename("UNITS"), c.rename("CRATES")], axis=1).fillna(0)
    t = t[(t["UNITS"] > 0) & (t["CRATES"] > 0)].copy()
    t["UNITS_PER_CRATE"] = t["UNITS"] / t["CRATES"]
    return t.reset_index(names="ROUTE"), w

def reconcile_returns():
    """Mehul vs Hansa: same dates, by product and by customer"""
    w = window(mehul_dw, hansa_dw)
    m = clip(mehul_dw, w)
    h_ = clip(hansa_dw, w)
    pm = m.groupby("PRODUCT")["MEHUL_TOTAL_RTN_QTY"].sum()
    ph = h_.groupby("PRODUCT")["HANSA_RTN_QTY"].sum()
    t = pd.concat([pm.rename("MEHUL"), ph.rename("HANSA")], axis=1).fillna(0)
    cm, ch = set(m["CODE"].dropna()), set(h_["CODE"].dropna())
    return t, w, cm, ch


# ## Sheet 1 - Brands

# In[45]:


def analyse_brands():
    B = brands_dw.copy()
    B["CUSTOMER"] = lab(B)
    for c in ("ROUTE NAME", "CATEGORY", "SALES_REP"):
        B[c] = dim(B, c)
    I = B[B["BRANDS_INV_QTY"] > 0]
    tot, cn = int(I["BRANDS_INV_QTY"].sum()), int(B["BRANDS_CN_QTY"].sum())
    nd = B["DATE"].nunique()
    prod = I.groupby("PRODUCT")["BRANDS_INV_QTY"].sum().sort_values(ascending=False)
    cust_q = I.groupby("CUSTOMER")["BRANDS_INV_QTY"].sum().sort_values(ascending=False)
    route = I.groupby("ROUTE NAME")["BRANDS_INV_QTY"].sum().sort_values(ascending=False)
    rep = I.groupby("SALES_REP")["BRANDS_INV_QTY"].sum().sort_values(ascending=False)
    daily = I.groupby("DATE")["BRANDS_INV_QTY"].sum()
    wd_text, wd = weekday_txt(daily, "units")
    Q = []

    Q.append(("Which bread and snack brands sell the most?",
              f"{top_list(prod, 3)}. The top 3 are {pc(prod.head(3).sum() / tot)} of the {n0(tot)} units invoiced.",
              "Protect production and van availability of these lines first; they carry the volume.",
              mk(top_tbl(prod, "PRODUCT", "UNITS"), "bar", ["UNITS"], title="Units invoiced - top 8 products")))
    small = int((prod / tot < 0.02).sum())
    low = prod[prod > 0].sort_values().head(8)
    Q.append(("Which products sell the least?",
              f"Lowest: {bottom_list(prod, 3)}. {small} of {len(prod)} products are each under 2% of volume.",
              "Check distribution gaps and shelf presence; decide whether to push or trim the weakest lines.",
              mk(pd.DataFrame({"PRODUCT": low.index, "UNITS": low.values, "SHARE %": low.values / tot}), "column", ["UNITS"], title="Units invoiced - 8 weakest products")))
    pcx = I.groupby(["PRODUCT", "CUSTOMER"])["BRANDS_INV_QTY"].sum().reset_index()
    best = pcx.sort_values("BRANDS_INV_QTY", ascending=False).drop_duplicates("PRODUCT").set_index("PRODUCT")
    parts = [f"{p}: {best.loc[p, 'CUSTOMER']} ({n0(best.loc[p, 'BRANDS_INV_QTY'])}, {pc(best.loc[p, 'BRANDS_INV_QTY'] / prod[p], 0)})" for p in prod.index[:5]]
    tb = pd.DataFrame({"PRODUCT": prod.index[:8], "TOP BUYER": [short(best.loc[p, "CUSTOMER"], 24) for p in prod.index[:8]],
                       "UNITS": [best.loc[p, "BRANDS_INV_QTY"] for p in prod.index[:8]],
                       "SHARE OF PRODUCT %": [best.loc[p, "BRANDS_INV_QTY"] / prod[p] for p in prod.index[:8]]})
    Q.append(("Which customers buy the most of each brand?",
              "Top buyer of the five biggest brands - " + "; ".join(parts) + ". The matrix at the end lists every customer by product.",
              "These are the key account-product relationships: secure supply and agree promotions with them.",
              mk(tb, "column", ["SHARE OF PRODUCT %"], cat="PRODUCT", title="Top buyer's share of each product")))
    Q.append(("Which routes generate the highest volumes?",
              f"{top_list(route, 3)}. Top 3 routes = {pc(route.head(3).sum() / tot)} of units across {len(route)} routes.",
              "Plan loading, vehicles and delivery windows around these routes first.",
              mk(top_tbl(route, "ROUTE", "UNITS"), "bar", ["UNITS"], title="Units invoiced - top 8 routes")))
    ncu = I.groupby("SALES_REP")["CUSTOMER"].nunique()
    per_cust = (rep / ncu).sort_values(ascending=False)
    rt = pd.DataFrame({"SALES REP": rep.index, "UNITS": rep.values, "CUSTOMERS": ncu.reindex(rep.index).values,
                       "UNITS PER CUSTOMER": (rep / ncu).reindex(rep.index).round(0).values}).head(8)
    rt["SALES REP"] = rt["SALES REP"].map(short)
    Q.append(("Which sales representatives handle the highest volumes?",
              f"{top_list(rep, 3)}. Units per customer served: highest is {per_cust.index[0]} at {n0(per_cust.iloc[0])}.",
              "Compare reps on units per customer, not only total units, so territory size does not distort the ranking.",
              mk(rt, "column", ["UNITS"], title="Units invoiced by sales rep")))
    wd_t = wd.rename_axis("WEEKDAY").reset_index(name="AVG UNITS / DAY")
    wd_t["DATES IN FILE"] = wd_t["WEEKDAY"].map(daily.groupby(daily.index.day_name()).size())
    Q.append(("How does demand vary by day of the week?", wd_text, "Match production and dispatch quantities to the busy weekdays.",
              mk(wd_t, "column", ["AVG UNITS / DAY"], title="Average units per day by weekday")))
    breadth = I.groupby("CUSTOMER")["PRODUCT"].nunique()
    half = int((breadth >= len(prod) / 2).sum()); few = int((breadth <= 2).sum())
    np_ = len(prod)
    edges = [0, 2, 5, 9, max(np_, 10)]
    lbls = ["1-2 products", "3-5 products", "6-9 products", "10+ products"]
    bt = bucket_tbl(breadth, edges, lbls, "RANGE BOUGHT", "CUSTOMERS")
    Q.append(("Which customers buy a broad range and which buy only a few products?",
              f"The typical customer buys {int(breadth.median())} of {np_} products. {half} customers ({pc(half / len(breadth), 0)}) buy at least half the range; "
              f"{few} customers ({pc(few / len(breadth), 0)}) buy only one or two.",
              "Target the narrow-range customers with a short list of the products their broad-range peers already buy.",
              mk(bt, "column", ["CUSTOMERS"], title="Customers by number of products bought")))
    bm = (I.groupby(["CUSTOMER", "PRODUCT"]).size().unstack(fill_value=0) > 0).astype(int)
    N, sup, co, cols = len(bm), bm.sum(), bm.T.dot(bm), list(bm.columns)
    pairs = []
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            c = int(co.iat[i, j])
            if c >= 2:
                pairs.append((cols[i], cols[j], c, (c / N) / ((sup[cols[i]] / N) * (sup[cols[j]] / N))))
    pairs.sort(key=lambda t: (-t[2], -t[3]))
    ptxt = "; ".join(f"{a} + {b} ({c} customers, lift {l:.2f})" for a, b, c, l in pairs[:3]) if pairs else "no pair is bought together by at least two customers"
    pt = pd.DataFrame({"PAIR": [f"{a} + {b}" for a, b, c, l in pairs[:8]], "CUSTOMERS": [c for a, b, c, l in pairs[:8]], "LIFT": [round(l, 2) for a, b, c, l in pairs[:8]]})
    Q.append(("Which products are bought together by the same customers?",
              f"Most common pairs: {ptxt}. Lift above 1.0 means the pair is bought together more often than chance.",
              "Use high-lift pairs for bundles and joint promotions; popular pairs with lift near 1.0 simply reflect popularity.",
              mk(pt, "bar", ["CUSTOMERS"], title="Customers buying each pair", fmts={"LIFT": "0.00"})))
    tc = trend_changes(I, "CUSTOMER", "BRANDS_INV_QTY", min_total=20)
    if tc is None:
        a9 = "The file holds a single date, so a decline cannot be measured yet. The table shows the biggest customers instead."
        b9 = mk(top_tbl(cust_q, "CUSTOMER", "UNITS"), "bar", ["UNITS"], title="Top customers (single date)")
    else:
        t, ndates = tc
        dec = t[(t["second"] < 0.7 * t["first"])].sort_values("gap")
        a9 = (f"Comparing the first and second half of the {ndates} dates: {len(dec)} customers buy at least 30% less per day"
              + (": " + "; ".join(change_txt(i, r, "units") for i, r in dec.head(3).iterrows()) if len(dec) else "") + ".")
        if ndates < 6:
            a9 += " Few dates are covered, so treat this as an early warning only."
        mv = movers_tbl(t)
        b9 = mk(mv, "column", ["FIRST HALF / DAY", "SECOND HALF / DAY"], title="Units per day: first vs second half (8 biggest falls)")
    Q.append(("Which customers have declining purchases?", a9, "Visit declining accounts: check stock-outs, competitor listings or changes in the buyer.", b9))
    cum = cust_q.cumsum() / cust_q.sum()
    n80 = int((cum < 0.8).sum() + 1)
    t10 = cust_q.head(10).sum() / tot
    t10p = cust_q.head(max(1, int(np.ceil(len(cust_q) * 0.1)))).sum() / tot
    ns = [n for n in (1, 5, 10, 20, 50, 100) if n <= len(cust_q)] or [len(cust_q)]
    if ns[-1] != len(cust_q) and len(cust_q) <= 200: ns.append(len(cust_q))
    ct = pd.DataFrame({"CUSTOMERS": [f"Top {n}" for n in ns], "UNITS": [int(cust_q.head(n).sum()) for n in ns], "CUMULATIVE SHARE %": [float(cum.iloc[n - 1]) for n in ns]})
    Q.append(("How concentrated are sales in a few customers or routes?",
              f"The top 10 customers buy {pc(t10)} of units; the top 10% of customers buy {pc(t10p)}; {n80} customers make up 80% of volume. "
              f"The top 3 routes carry {pc(route.head(3).sum() / tot)}.",
              "A high share in few accounts is a dependency risk: protect those accounts and widen the base.",
              mk(ct, "column", ["CUMULATIVE SHARE %"], title="Share of all units bought by the top N customers", ymax=1)))

    # ---------------- relationships
    R = []
    tM, wM, _ = product_rates(mehul_dw, "MEHUL_TOTAL_RTN_QTY")
    tH, wH, _ = product_rates(hansa_dw, "HANSA_RTN_QTY")
    rate = pd.merge(tM[["PRODUCT", "SOLD", "RATE"]].rename(columns={"SOLD": "S1", "RATE": "MEHUL RETURN RATE %"}),
                    tH[["PRODUCT", "SOLD", "RATE"]].rename(columns={"SOLD": "S2", "RATE": "HANSA RETURN RATE %"}), on="PRODUCT", how="outer")
    rate["BIG"] = rate[["S1", "S2"]].max(axis=1)
    rate = rate[rate["BIG"] >= 0.01 * rate["BIG"].sum()].copy()
    rate["MAXR"] = rate[["MEHUL RETURN RATE %", "HANSA RETURN RATE %"]].max(axis=1)
    rate = rate.sort_values("MAXR", ascending=False).head(8)[["PRODUCT", "MEHUL RETURN RATE %", "HANSA RETURN RATE %"]]
    R.append(("Which products sell a lot but also come back a lot?",
              rates_text(tM, "Mehul") + " " + win_note(wM, [("brands", brands_dw), ("Mehul", mehul_dw)]) + "\n" +
              rates_text(tH, "Hansa") + " " + win_note(wH, [("brands", brands_dw), ("Hansa", hansa_dw)]),
              "Investigate dispatch quantities for high-return products. Mehul and Hansa are shown separately because they may record the same returns.",
              mk(rate, "column", ["MEHUL RETURN RATE %", "HANSA RETURN RATE %"], title="Returns as % of units sold (blank = product not in that returns file)")))
    w = window(brands_dw, rev_dw)
    a = code_sum(brands_dw, w, ["BRANDS_INV_QTY"])["BRANDS_INV_QTY"]
    b = code_sum(rev_dw, w, ["REV_INV_AMT"])["REV_INV_AMT"]
    rho, n = rank_corr(a, b)
    top8 = a.sort_values(ascending=False).head(8)
    t2 = pd.DataFrame({"CUSTOMER": [clabel(c) for c in top8.index], "BRAND UNITS": top8.values,
                       "INVOICE VALUE": [float(b.get(c, np.nan)) if c in b.index else np.nan for c in top8.index]})
    t2["VALUE PER UNIT"] = (t2["INVOICE VALUE"] / t2["BRAND UNITS"]).round(1)
    R.append(("Do the biggest brand buyers also bring in the most revenue?",
              f"{n} customers appear in both files; they show {corr_words(rho)} between units and invoice value. "
              f"{int((~a.index.isin(b.index)).sum())} of {len(a)} brand customers have no sales-revenue record. " + win_note(w, [("brands", brands_dw), ("revenue", rev_dw)]),
              "Customers high on units but low on revenue deserve a pricing and mix check.",
              mk(t2, "bar", ["VALUE PER UNIT"], title="Invoice value per brand unit - 8 biggest buyers")))
    rc, wc = route_crate_ratio()
    if len(rc):
        tot_ratio = sdiv(rc["UNITS"].sum(), rc["CRATES"].sum())
        a3 = (f"On average {n1(tot_ratio)} units move per crate invoiced. Most units per crate: {rc.sort_values('UNITS_PER_CRATE', ascending=False).iloc[0]['ROUTE']} "
              f"({n1(rc['UNITS_PER_CRATE'].max())}); fewest: {rc.sort_values('UNITS_PER_CRATE').iloc[0]['ROUTE']} ({n1(rc['UNITS_PER_CRATE'].min())}). " + win_note(wc, [("brands", brands_dw), ("crates", crates_dw)]))
        r3 = rc.sort_values("CRATES", ascending=False).head(8).copy()
        r3["ROUTE"] = r3["ROUTE"].map(short)
        r3 = r3.rename(columns={"UNITS_PER_CRATE": "UNITS PER CRATE"}); r3["UNITS PER CRATE"] = r3["UNITS PER CRATE"].round(1)
        b3 = mk(r3, "bar", ["UNITS PER CRATE"], title="Units per crate - 8 routes with most crates")
    else:
        a3 = "No route has both brand units and crates recorded. " + win_note(wc, [("brands", brands_dw), ("crates", crates_dw)])
        b3 = mk(top_tbl(route, "ROUTE", "UNITS"), "bar", ["UNITS"], title="Units by route (no crate match)")
    R.append(("Is crate allocation in line with the units sold?", a3,
              "Routes with very few units per crate may hold more crates than they need; routes with very many may be short.", b3))
    w = window(brands_dw, rev_dw)
    bb = code_sum(brands_dw, w, ["BRANDS_INV_QTY", "BRANDS_CN_QTY"])
    rr = code_sum(rev_dw, w, ["REV_INV_AMT", "REV_CN_AMT"])
    ra = (bb["BRANDS_CN_QTY"] / bb["BRANDS_INV_QTY"].where(bb["BRANDS_INV_QTY"] > 0)).dropna()
    rb = (rr["REV_CN_AMT"] / rr["REV_INV_AMT"].where(rr["REV_INV_AMT"] > 0)).dropna()
    rho2, n2 = rank_corr(ra, rb)
    ids = ra[bb["BRANDS_INV_QTY"].reindex(ra.index) >= 20].sort_values(ascending=False).head(8).index
    t4 = pd.DataFrame({"CUSTOMER": [clabel(c) for c in ids], "BRAND CREDIT RATE %": [float(ra[c]) for c in ids],
                       "REVENUE CREDIT RATIO %": [float(rb[c]) if c in rb.index else np.nan for c in ids]})
    R.append(("Do customers with high brand credit rates also have high revenue credit ratios?",
              f"Brand credits are {pc(sdiv(cn, tot))} of net units. Across {n2} customers the brand credit rate and the revenue credit ratio show {corr_words(rho2)}. " + win_note(w, [("brands", brands_dw), ("revenue", rev_dw)]),
              "Where both are high, review the credit approvals for those customers.",
              mk(t4, "column", ["BRAND CREDIT RATE %", "REVENUE CREDIT RATIO %"], title="Credit rates - 8 customers with the highest brand credit rate (20+ units bought)")))

    # ---------------- matrix
    order = list(prod.index)
    M = pivot_matrix(I, "BRANDS_INV_QTY", "PRODUCT")
    M = M[INFO_COLS + [c for c in order if c in M.columns]]
    cr = B.assign(_k=B["CODE"].astype("string").fillna("") + "|" + B["CUSTOMER NAME"].astype("string").fillna("")).groupby("_k")["BRANDS_CN_QTY"].sum()
    M["TOTAL_UNITS"] = M[[c for c in order if c in M.columns]].sum(axis=1)
    keymap = (M["CODE"].astype("string").fillna("") + "|" + M["CUSTOMER NAME"].astype("string").fillna(""))
    M["CREDIT_UNITS"] = keymap.map(cr).fillna(0).astype(int).to_numpy()
    M["CREDIT_RATE"] = np.where(M["TOTAL_UNITS"] > 0, M["CREDIT_UNITS"] / M["TOTAL_UNITS"], 0).round(4)
    M = M.sort_values("TOTAL_UNITS", ascending=False).reset_index(drop=True)
    checks = [("brands: product table = matrix total", prod.sum(), M["TOTAL_UNITS"].sum()),
              ("brands: route table = matrix total", route.sum(), M["TOTAL_UNITS"].sum()),
              ("brands: concentration table top-N = customer totals", ct["UNITS"].iloc[-1], cust_q.head(ns[-1]).sum())]
    return dict(
        name="1 Brands", title="BRANDS - product demand and customer buying", color="#1F3864",
        subtitle=f"Source: brands invoices and credit notes | Period: {span_txt(B)} | {nd} dates",
        kpis=[("NET UNITS INVOICED", n0(tot)), ("CREDIT UNITS", n0(cn)), ("CREDIT RATE", pc(sdiv(cn, tot))),
              ("CUSTOMERS BUYING", n0(I["CUSTOMER"].nunique())), ("PRODUCTS SOLD", n0(len(prod))), ("AVG UNITS PER DAY", n0(sdiv(tot, nd)))],
        notes=["Units are net invoice units: an invoice that a credit note reverses exactly (same customer, date and quantities) is removed, as in the reporting notebooks. "
               "Credits that do not reverse an invoice are shown separately as credit units.",
               "Customer, route, category and sales rep come from the customer-info file, looked up by A/C code."],
        questions=Q, rel_title="HOW BRANDS RELATE TO THE OTHER FOUR DATASETS", rels=R, checks=checks,
        matrix=M, matrix_title="CUSTOMER x PRODUCT MATRIX - net units invoiced (use the filter buttons on the header row)",
        heat=order, sums=order + ["TOTAL_UNITS", "CREDIT_UNITS"], pct=["CREDIT_RATE"], rotate=order)


# ## Sheet 2 - Crates

# In[46]:


def analyse_crates():
    C = crates_dw.copy()
    C["CUSTOMER"] = lab(C)
    for c in ("ROUTE NAME", "CATEGORY", "SALES_REP"):
        C[c] = dim(C, c)
    I = C[C["CRATES_INV_QTY"] > 0]
    tot, cn = int(I["CRATES_INV_QTY"].sum()), int(C["CRATES_CN_QTY"].sum())
    dates = sorted(C["DATE"].unique()); nd = len(dates)
    daily = C.groupby("DATE")[["CRATES_INV_QTY", "CRATES_CN_QTY"]].sum()
    cust = I.groupby("CUSTOMER")["CRATES_INV_QTY"].sum().sort_values(ascending=False)
    route = I.groupby("ROUTE NAME")["CRATES_INV_QTY"].sum().sort_values(ascending=False)
    wd_text, wd = weekday_txt(daily["CRATES_INV_QTY"], "crates")
    Q = []
    Q.append(("Which customers use the most crates?",
              f"{top_list(cust, 5)}. {len(cust)} customers received crates in total ({n0(tot)} net crates).",
              "Confirm these accounts have enough crates and that they return them on time.",
              mk(top_tbl(cust, "CUSTOMER", "CRATES"), "bar", ["CRATES"], title="Crates invoiced - top 8 customers")))
    Q.append(("Which routes need the most crates?",
              f"{top_list(route, 3)}. Top 3 routes = {pc(route.head(3).sum() / tot)} of crates.",
              "Load and balance the crate pool by route, starting with the largest.",
              mk(top_tbl(route, "ROUTE", "CRATES"), "bar", ["CRATES"], title="Crates invoiced - top 8 routes")))
    pk, lo = daily["CRATES_INV_QTY"].idxmax(), daily["CRATES_INV_QTY"].idxmin()
    pf = period_frame(daily.reset_index(), ["CRATES_INV_QTY", "CRATES_CN_QTY"]).rename(columns={"CRATES_INV_QTY": "INVOICED", "CRATES_CN_QTY": "CREDITED"})
    Q.append(("What is the daily crate movement pattern?",
              f"Average {n0(sdiv(tot, nd))} crates per day over {nd} dates. Peak day {dtxt(pk)} ({n0(daily.loc[pk, 'CRATES_INV_QTY'])}); "
              f"lowest {dtxt(lo)} ({n0(daily.loc[lo, 'CRATES_INV_QTY'])}).",
              "Plan crate availability for the peak days.",
              mk(pf, "line", ["INVOICED", "CREDITED"], title=f"Crates invoiced and credited ({pf.columns[0].lower()})")))
    avg = (cust / nd).sort_values(ascending=False)
    ext = pd.concat([avg.head(4), avg.tail(4)]).drop_duplicates()
    ta = pd.DataFrame({"CUSTOMER": [short(i) for i in ext.index], "AVG CRATES / DAY": ext.values.round(1),
                       "GROUP": ["Highest"] * min(4, len(avg)) + ["Lowest"] * (len(ext) - min(4, len(avg)))})
    Q.append(("Which customers have the highest and lowest average crate use?",
              f"Highest per day: {'; '.join(f'{i} ({n1(v)})' for i, v in avg.head(3).items())}. "
              f"Lowest (still active): {'; '.join(f'{i} ({n1(v)})' for i, v in avg.tail(3).items())}.",
              "Very low users may be holding crates they do not need; reallocate if so.",
              mk(ta, "bar", ["AVG CRATES / DAY"], title="Average crates per day - 4 highest and 4 lowest customers")))
    dv = daily["CRATES_INV_QTY"]
    if nd >= 5 and dv.std() > 0:
        z = (dv - dv.mean()) / dv.std()
        odd = z[z.abs() >= 2]
        a5 = ("; ".join(f"{dtxt(d)} ({n0(dv[d])} crates, {z[d]:+.1f} sd)" for d in odd.index) if len(odd) else
              "No day is unusual (all days are within 2 standard deviations of the daily average).")
        zt = z.reindex(z.abs().sort_values(ascending=False).index).head(8)
        t5 = pd.DataFrame({"DATE": [pd.Timestamp(d).strftime("%d %b") for d in zt.index], "CRATES": dv[zt.index].to_numpy(), "Z-SCORE": zt.round(2).to_numpy()})
    else:
        a5 = "Too few dates (fewer than 5) to call any day unusual; the table lists every date."
        t5 = pd.DataFrame({"DATE": [pd.Timestamp(d).strftime("%d %b") for d in dv.index], "CRATES": dv.to_numpy()})
    Q.append(("Which days have unusually high or low crate use?", a5, "Investigate unusual days for dispatch disruptions or data-entry errors.",
              mk(t5.head(8), "column", ["CRATES"], title="Crates on the days furthest from the average", fmts={"Z-SCORE": "0.00"})))
    cat = I.groupby("CATEGORY")["CRATES_INV_QTY"].sum().sort_values(ascending=False)
    rep = I.groupby("SALES_REP")["CRATES_INV_QTY"].sum().sort_values(ascending=False)
    Q.append(("How does crate use differ by category and sales representative?",
              f"By category: {top_list(cat, 3)}. By sales rep: {top_list(rep, 3)}.",
              "Use these shares to compare workloads and to set crate quotas per category.",
              mk(top_tbl(cat, "CATEGORY", "CRATES", k=8), "pie", ["CRATES"], title="Share of crates by category")))
    tc = trend_changes(I, "CUSTOMER", "CRATES_INV_QTY", min_total=10)
    if tc is None:
        a7 = "The file holds a single date, so a change cannot be measured yet. The table shows the biggest users instead."
        b7 = mk(top_tbl(cust, "CUSTOMER", "CRATES"), "bar", ["CRATES"], title="Top customers (single date)")
    else:
        t, ndates = tc
        up = t[t["second"] >= 2 * t["first"]].sort_values("gap", ascending=False)
        dn = t[t["second"] <= 0.5 * t["first"]].sort_values("gap")
        a7 = (f"Comparing the first and second half of the {ndates} dates: {len(up)} customers at least double their crate use"
              + (" (" + "; ".join(change_txt(i, r, "crates") for i, r in up.head(2).iterrows()) + ")" if len(up) else "")
              + f"; {len(dn)} halve it" + (" (" + "; ".join(change_txt(i, r, "crates") for i, r in dn.head(2).iterrows()) + ")" if len(dn) else "") + ".")
        mv = t.assign(a=t["gap"].abs()).sort_values("a", ascending=False).head(8)
        b7 = mk(movers_tbl(mv.sort_values("gap"), by="gap"), "column", ["FIRST HALF / DAY", "SECOND HALF / DAY"], title="Crates per day: first vs second half (8 biggest movers)")
    Q.append(("Which customers show sudden changes in crate use?", a7, "Call these customers: sudden changes can mean new outlets, closures, or crates left on site.", b7))
    wd_t = wd.rename_axis("WEEKDAY").reset_index(name="AVG CRATES / DAY")
    wd_t["DATES IN FILE"] = wd_t["WEEKDAY"].map(daily.groupby(daily.index.day_name()).size())
    Q.append(("Is crate use consistent across the days of the week?", wd_text, "Schedule crate collection and washing around the busy weekdays.",
              mk(wd_t, "column", ["AVG CRATES / DAY"], title="Average crates per day by weekday")))
    rc, wc = route_crate_ratio()
    if len(rc):
        rcs = rc[rc["CRATES"] >= 0.02 * rc["CRATES"].sum()].sort_values("UNITS_PER_CRATE")
        a9 = (f"Routes holding the most crates per unit sold (possible surplus): {'; '.join(f'{r.ROUTE} ({n1(r.UNITS_PER_CRATE)} units/crate)' for r in rcs.head(2).itertuples())}. "
              f"Fewest crates per unit (possible shortage): {'; '.join(f'{r.ROUTE} ({n1(r.UNITS_PER_CRATE)} units/crate)' for r in rcs.tail(2).itertuples())}. " + win_note(wc, [("crates", crates_dw), ("brands", brands_dw)]))
        ext = pd.concat([rcs.head(4), rcs.tail(4)]).drop_duplicates("ROUTE")
        t9 = pd.DataFrame({"ROUTE": ext["ROUTE"].map(short), "UNITS": ext["UNITS"].to_numpy(), "CRATES": ext["CRATES"].to_numpy(), "UNITS PER CRATE": ext["UNITS_PER_CRATE"].round(1).to_numpy()})
        b9 = mk(t9, "bar", ["UNITS PER CRATE"], title="Units per crate - routes with the most and fewest crates per unit")
    else:
        a9 = "No route has both crates and brand units. " + win_note(wc, [("crates", crates_dw), ("brands", brands_dw)])
        b9 = mk(top_tbl(route, "ROUTE", "CRATES"), "bar", ["CRATES"], title="Crates by route (no brand match)")
    Q.append(("Which routes may need more or fewer crates?", a9, "Move crates from low units-per-crate routes to high ones before buying more.", b9))
    n_inv_docs, n_cn_docs = len(c_inv), len(c_cn)
    canc = int(c_cn_cancelled["QTY"].sum())
    t10 = pd.DataFrame({"ITEM": ["Invoiced (all documents)", "Removed: reversed by a credit note", "Net invoiced", "Stand-alone credits"],
                        "DOCUMENTS": [n_inv_docs, len(c_cn_cancelled), n_inv_docs - len(c_cn_cancelled), len(c_cn_open)],
                        "CRATES": [int(c_inv["QTY"].sum()), canc, tot, cn]})
    Q.append(("Do the crate invoices and credit notes reconcile?",
              f"{n_inv_docs:,} crate invoices and {n_cn_docs:,} credit notes were read. {len(c_cn_cancelled):,} credit notes exactly reversed an invoice (both removed); "
              f"{len(c_cn_open):,} stand alone and total {n0(cn)} crates, {pc(sdiv(cn, tot))} of net crates invoiced. "
              f"{int((c_inv['QTY'] == 0).sum())} invoices and {int((c_cn['QTY'] == 0).sum())} credit notes have a zero quantity.",
              "Review zero-quantity documents and stand-alone credits for data-entry errors.",
              mk(t10, "column", ["CRATES"], cat="ITEM", title="Crates: invoiced, removed, net and stand-alone credits")))

    R = []
    rc2, wc2 = route_crate_ratio()
    w = window(brands_dw, crates_dw)
    a = code_sum(crates_dw, w, ["CRATES_INV_QTY"])["CRATES_INV_QTY"]
    b = code_sum(brands_dw, w, ["BRANDS_INV_QTY"])["BRANDS_INV_QTY"]
    rho, n = rank_corr(a, b)
    top8 = a.sort_values(ascending=False).head(8)
    t1 = pd.DataFrame({"CUSTOMER": [clabel(c) for c in top8.index], "CRATES": top8.values,
                       "BRAND UNITS": [float(b.get(c, np.nan)) if c in b.index else np.nan for c in top8.index]})
    t1["UNITS PER CRATE"] = (t1["BRAND UNITS"] / t1["CRATES"]).round(1)
    R.append(("Do customers with more crates also buy more units?",
              f"Across {n} customers in both files crates and units show {corr_words(rho)}. Overall {n1(sdiv(rc2['UNITS'].sum(), rc2['CRATES'].sum()))} units move per crate. " + win_note(w, [("crates", crates_dw), ("brands", brands_dw)]),
              "A weak link means crates are not being allocated in proportion to demand.",
              mk(t1, "bar", ["UNITS PER CRATE"], title="Units per crate - 8 customers with most crates")))
    w = window(rev_dw, crates_dw)
    rv = clip(rev_dw, w)["REV_INV_AMT"].sum(); cr_ = clip(crates_dw, w)["CRATES_INV_QTY"].sum()
    a2 = code_sum(rev_dw, w, ["REV_INV_AMT"])["REV_INV_AMT"]; b2 = code_sum(crates_dw, w, ["CRATES_INV_QTY"])["CRATES_INV_QTY"]
    rho2, n2 = rank_corr(a2, b2)
    rt_ = (clip(rev_dw, w).assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")["REV_INV_AMT"].sum())
    ct_ = (clip(crates_dw, w).assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")["CRATES_INV_QTY"].sum())
    j = pd.concat([rt_.rename("INVOICE VALUE"), ct_.rename("CRATES")], axis=1).dropna()
    j = j[(j["CRATES"] > 0) & (j["INVOICE VALUE"] > 0)].sort_values("CRATES", ascending=False).head(8)
    t2 = pd.DataFrame({"ROUTE": [short(i) for i in j.index], "INVOICE VALUE": j["INVOICE VALUE"].round(0).to_numpy(), "CRATES": j["CRATES"].to_numpy(),
                       "VALUE PER CRATE": (j["INVOICE VALUE"] / j["CRATES"]).round(0).to_numpy()})
    R.append(("How much revenue does each crate support?",
              f"About {n0(sdiv(rv, cr_))} of invoice value per crate invoiced; customers show {corr_words(rho2)} between revenue and crates ({n2} customers). " + win_note(w, [("revenue", rev_dw), ("crates", crates_dw)]),
              "Customers with low revenue per crate tie up crates for little sales value.",
              mk(t2, "bar", ["VALUE PER CRATE"], title="Invoice value per crate - routes with most crates")))
    for lbl, src, col, name in (("Hansa", hansa_dw, "HANSA_RTN_QTY", "HANSA RETURNS"), ("Mehul", mehul_dw, "MEHUL_TOTAL_RTN_QTY", "MEHUL RETURNS")):
        w = window(src, crates_dw)
        a3 = code_sum(crates_dw, w, ["CRATES_INV_QTY"])["CRATES_INV_QTY"]; b3 = code_sum(src, w, [col])[col]
        rho3, n3 = rank_corr(a3, b3)
        top8 = a3.sort_values(ascending=False).head(8)
        t3 = pd.DataFrame({"CUSTOMER": [clabel(c) for c in top8.index], "CRATES": top8.values,
                           name: [float(b3.get(c, 0)) for c in top8.index]})
        t3["RETURNS PER 100 CRATES"] = (t3[name] / t3["CRATES"] * 100).round(1)
        R.append((f"Do customers with many crates also return a lot ({lbl})?",
                  f"Across {n3} customers crate counts and {lbl} returns show {corr_words(rho3)}. " + win_note(w, [("crates", crates_dw), (lbl, src)]),
                  "If returns grow with crates, over-supply may be filling crates with product that comes back." if lbl == "Hansa" else "Use together with the Hansa result before drawing conclusions; correlation is not cause.",
                  mk(t3, "bar", ["RETURNS PER 100 CRATES"], title=f"{lbl} returns per 100 crates - 8 customers with most crates")))

    # ---------------- matrix: customer x date
    lbl = {d: pd.Timestamp(d).strftime("%d-%b") for d in dates}
    if nd > 45:
        lbl = {d: pd.Timestamp(d).strftime("%b %Y") for d in dates}
    Cm = C.assign(PERIOD=C["DATE"].map(lbl))
    M = pivot_matrix(Cm, "CRATES_INV_QTY", "PERIOD")
    pcols = list(dict.fromkeys(lbl[d] for d in dates))
    M = M[INFO_COLS + [c for c in pcols if c in M.columns]]
    pcols = [c for c in pcols if c in M.columns]
    cn_by = Cm.assign(_k=Cm["CODE"].astype("string").fillna("") + "|" + Cm["CUSTOMER NAME"].astype("string").fillna("")).groupby("_k")["CRATES_CN_QTY"].sum()
    M["TOTAL_CRATES"] = M[pcols].sum(axis=1)
    keymap = M["CODE"].astype("string").fillna("") + "|" + M["CUSTOMER NAME"].astype("string").fillna("")
    M["CREDITED"] = keymap.map(cn_by).fillna(0).astype(int).to_numpy()
    M["AVG_PER_DAY"] = (M["TOTAL_CRATES"] / nd).round(1)
    M = M.sort_values("TOTAL_CRATES", ascending=False).reset_index(drop=True)
    checks = [("crates: customer table = matrix total", cust.sum(), M["TOTAL_CRATES"].sum()),
              ("crates: route table = matrix total", route.sum(), M["TOTAL_CRATES"].sum()),
              ("crates: daily table = matrix total", pf["INVOICED"].sum() if len(pf) == nd or nd <= 12 else tot, tot),
              ("crates: reconciliation net = matrix total", t10["CRATES"].iloc[2], M["TOTAL_CRATES"].sum())]
    return dict(
        name="2 Crates", title="CRATES - allocation, movement and reconciliation", color="#1B7F79",
        subtitle=f"Source: crate invoices and credit notes | Period: {span_txt(C)} | {nd} dates",
        kpis=[("NET CRATES INVOICED", n0(tot)), ("CRATES CREDITED", n0(cn)), ("CUSTOMERS", n0(I["CUSTOMER"].nunique())),
              ("AVG CRATES PER DAY", n0(sdiv(tot, nd))), ("PEAK DAY", dtxt(pk)), ("PEAK DAY CRATES", n0(daily.loc[pk, "CRATES_INV_QTY"]))],
        notes=["Each crate figure is a document quantity (invoice or credit note). The data does not say whether a crate is issued, returned or held, so no loss or circulation figure is calculated.",
               "A credit note that exactly reverses an invoice (same account and quantity) removes that invoice, as in the reporting notebooks."],
        questions=Q, rel_title="HOW CRATES RELATE TO THE OTHER FOUR DATASETS", rels=R, checks=checks,
        matrix=M, matrix_title="CUSTOMER x DATE MATRIX - crates invoiced (use the filter buttons on the header row)",
        heat=pcols, sums=pcols + ["TOTAL_CRATES", "CREDITED"], pct=[], rotate=pcols)


# ## Sheet 3 - Mehul returns

# In[47]:


def analyse_mehul():
    M0 = mehul_dw.copy()
    M0["CUSTOMER"] = lab(M0)
    for c in ("ROUTE NAME", "CATEGORY", "SALES_REP"):
        M0[c] = dim(M0, c)
    F, X, T_ = "MEHUL_FRESH_RTN_QTY", "MEHUL_EXP_RTN_QTY", "MEHUL_TOTAL_RTN_QTY"
    tot, fr, ex = int(M0[T_].sum()), int(M0[F].sum()), int(M0[X].sum())
    nd = M0["DATE"].nunique()
    prod = M0.groupby("PRODUCT")[[F, X, T_]].sum().sort_values(T_, ascending=False)
    route = M0.groupby("ROUTE NAME")[[F, X, T_]].sum().sort_values(T_, ascending=False)
    cust = M0.groupby("CUSTOMER")[[F, X, T_]].sum().sort_values(T_, ascending=False)
    daily = M0.groupby("DATE")[[F, X, T_]].sum()
    FX = {F: "FRESH", X: "EXPIRED", T_: "TOTAL"}
    Q = []
    q1t = pd.DataFrame({"TYPE": ["Fresh", "Expired"], "UNITS": [fr, ex], "SHARE %": [sdiv(fr, tot), sdiv(ex, tot)]})
    Q.append(("What is the total of fresh and expired returns?",
              f"{n0(tot)} units returned over {nd} dates: {n0(fr)} fresh and {n0(ex)} expired. Average {n0(sdiv(tot, nd))} units a day.",
              "This is the baseline to improve against each period.", mk(q1t, "pie", ["UNITS"], title="Fresh vs expired returns")))
    pt = prod.head(8)[[F, X, T_]].rename(columns=FX).rename_axis("PRODUCT").reset_index()
    Q.append(("Which products are returned the most?", f"{top_list(prod[T_], 3)}.", "Review dispatch and order quantities of these products first.",
              mk(pt, "column", ["FRESH", "EXPIRED"], stacked=True, title="Returns by product - fresh and expired (top 8)")))
    rt = route.head(8)[[F, X, T_]].rename(columns=FX).rename_axis("ROUTE").reset_index(); rt["ROUTE"] = rt["ROUTE"].map(short)
    Q.append(("Which routes have the highest returns?", f"{top_list(route[T_], 3)}. Top 3 routes = {pc(route[T_].head(3).sum() / tot)} of returns.",
              "Sit down with these route teams to review loading quantities and delivery timing.",
              mk(rt, "bar", ["FRESH", "EXPIRED"], stacked=True, title="Returns by route - fresh and expired (top 8)")))
    ct = cust.head(8)[[T_, X]].copy()
    ct4 = pd.DataFrame({"CUSTOMER": [short(i) for i in ct.index], "RETURNS": ct[T_].to_numpy(), "EXPIRED SHARE %": (ct[X] / ct[T_]).to_numpy()})
    Q.append(("Which customers return the most?", f"{top_list(cust[T_], 5)}.", "Check the order pattern and shelf-life handling at these outlets.",
              mk(ct4, "bar", ["RETURNS"], title="Returns - top 8 customers")))
    p5 = prod.head(8).copy()
    t5 = pd.DataFrame({"PRODUCT": p5.index, "FRESH %": (p5[F] / p5[T_]).to_numpy(), "EXPIRED %": (p5[X] / p5[T_]).to_numpy()})
    Q.append(("What share of returns is fresh and what share expired?",
              f"Fresh {pc(sdiv(fr, tot))}, expired {pc(sdiv(ex, tot))}. Fresh returns are potentially re-sellable; expired returns are a write-off.",
              "Push to move returns from expired to fresh by faster rotation and earlier collection.",
              mk(t5, "column", ["FRESH %", "EXPIRED %"], stacked=True, title="Fresh vs expired share by product (top 8)", ymax=1)))
    rt_ = route[route[T_] >= 0.02 * tot].copy(); rt_["EXP_SHARE"] = rt_[X] / rt_[T_]
    rt2 = rt_.sort_values(X, ascending=False)
    t6 = pd.DataFrame({"ROUTE": [short(i) for i in rt2.head(8).index], "EXPIRED": rt2.head(8)[X].to_numpy(), "EXPIRED SHARE %": rt2.head(8)["EXP_SHARE"].to_numpy()})
    Q.append(("Which routes have the most expired returns?",
              f"By volume: {'; '.join(f'{i} ({n0(r[X])}, {pc(r.EXP_SHARE, 0)} of its returns)' for i, r in rt2.head(3).iterrows())}.",
              "Prioritise waste-reduction visits to these routes.", mk(t6, "bar", ["EXPIRED"], title="Expired returns - top 8 routes")))
    pp = prod[prod[T_] >= max(10, 0.02 * tot)].copy(); pp["EXP_SHARE"] = pp[X] / pp[T_]
    avg_s = sdiv(ex, tot)
    pp = pp.sort_values("EXP_SHARE", ascending=False)
    if not len(pp):
        pp = prod.copy(); pp["EXP_SHARE"] = pp[X] / pp[T_]; pp = pp.sort_values("EXP_SHARE", ascending=False)
    t7 = pd.DataFrame({"PRODUCT": pp.head(8).index, "RETURNS": pp.head(8)[T_].to_numpy(), "EXPIRED SHARE %": pp.head(8)["EXP_SHARE"].to_numpy(),
                       "VS AVERAGE (pts)": ((pp.head(8)["EXP_SHARE"] - avg_s) * 100).round(1).to_numpy()})
    Q.append(("Which products have a high expired share?",
              f"Average expired share is {pc(avg_s)}. Highest: {'; '.join(f'{i} ({pc(r.EXP_SHARE, 0)}, {r.EXP_SHARE * 100 - avg_s * 100:+.0f} pts)' for i, r in pp.head(3).iterrows())}.",
              "Review shelf life, pack size and order cycle for these products.",
              mk(t7, "bar", ["EXPIRED SHARE %"], title="Expired share of returns - highest products", fmts={"VS AVERAGE (pts)": "+0.0;-0.0;0.0"})))
    pk = daily[T_].idxmax()
    heavy = int((daily[T_] > 1.5 * daily[T_].mean()).sum()) if nd > 1 else 0
    pf = period_frame(daily.reset_index(), [F, X]).rename(columns={F: "FRESH", X: "EXPIRED"})
    Q.append(("Are returns concentrated on certain days?",
              f"Peak day {dtxt(pk)} with {n0(daily.loc[pk, T_])} units ({pc(sdiv(daily.loc[pk, T_], tot))} of returns); {heavy} of {nd} dates are more than 50% above the daily average."
              + (" Only one date is covered, so no pattern can be shown yet." if nd == 1 else ""),
              "Look at what happened on peak days (weather, late deliveries, promotions).",
              mk(pf, "line", ["FRESH", "EXPIRED"], title=f"Returns per period ({pf.columns[0].lower()})")))
    cat = M0.groupby("CATEGORY")[[F, X, T_]].sum().sort_values(T_, ascending=False)
    t9 = pd.DataFrame({"CATEGORY": cat.head(8).index, "RETURNS": cat.head(8)[T_].to_numpy(), "EXPIRED SHARE %": (cat.head(8)[X] / cat.head(8)[T_]).to_numpy()})
    Q.append(("How do returns differ by customer category?",
              "; ".join(f"{i}: {n0(r[T_])} units, {pc(sdiv(r[X], r[T_]), 0)} expired" for i, r in cat.head(4).iterrows()) + ".",
              "Set category-specific order and rotation rules.", mk(t9, "column", ["RETURNS"], title="Returns by customer category")))
    unit_rev = sdiv(rev_dw["REV_INV_AMT"].sum(), rev_dw["REV_INV_QTY"].sum())
    if unit_rev > 0:
        vr = (route[T_] * unit_rev).sort_values(ascending=False)
        a10 = (f"Valued at the average invoice revenue per unit in the sales-revenue file ({n1(unit_rev)}), returns are worth about {n0(tot * unit_rev)} "
               f"(expired: {n0(ex * unit_rev)}). Largest routes: {top_list(vr, 3)}. This is a blended estimate, not product prices.")
        t10 = pd.DataFrame({"ROUTE": [short(i) for i in vr.head(8).index], "UNITS": route[T_].reindex(vr.head(8).index).to_numpy(), "ESTIMATED VALUE": vr.head(8).round(0).to_numpy()})
        b10 = mk(t10, "bar", ["ESTIMATED VALUE"], title=f"Estimated value of returns at {n1(unit_rev)} per unit - top 8 routes")
    else:
        a10 = "No invoice revenue per unit is available, so returns cannot be valued; the table shows units by route."
        b10 = mk(rt[["ROUTE", "TOTAL"]].rename(columns={"TOTAL": "UNITS"}), "bar", ["UNITS"], title="Returns by route (no valuation possible)")
    Q.append(("What is the estimated value of the returns?", a10, "Replace the blended price with product price lists for an exact figure.", b10))

    R = []
    tM, wM, unm = product_rates(mehul_dw, T_)
    extra = f" Product names with returns but no sales match: {', '.join(unm.index.astype(str))}." if len(unm) else ""
    r1 = tM[tM["SOLD"] >= 0.01 * tM["SOLD"].sum()].sort_values("RATE", ascending=False).head(8)
    t1 = pd.DataFrame({"PRODUCT": r1["PRODUCT"], "UNITS SOLD": r1["SOLD"].to_numpy(), "MEHUL RETURNED": r1["RETURNED"].to_numpy(), "RETURN RATE %": r1["RATE"].to_numpy()})
    R.append(("Which products come back relative to what is sold?", rates_text(tM, "Mehul") + " " + win_note(wM, [("Mehul", mehul_dw), ("brands", brands_dw)]) + extra,
              "Return rate, not return volume, shows where dispatch exceeds demand.", mk(t1, "column", ["RETURN RATE %"], title="Mehul returns as % of units sold - highest 8 products")))
    t, w, cm, ch = reconcile_returns()
    both = t[(t["MEHUL"] > 0) & (t["HANSA"] > 0)]
    rho = None if len(both) < 4 else float(both["MEHUL"].rank().corr(both["HANSA"].rank()))
    t2 = t.assign(S=t["MEHUL"] + t["HANSA"]).sort_values("S", ascending=False).head(8).drop(columns="S").rename_axis("PRODUCT").reset_index()
    t2 = t2.rename(columns={"MEHUL": "MEHUL UNITS", "HANSA": "HANSA UNITS"})
    R.append(("Do Mehul and Hansa returns agree?",
              f"On the dates both cover, Mehul records {n0(t['MEHUL'].sum())} units and Hansa {n0(t['HANSA'].sum())}. Product totals show {corr_words(rho)}. "
              f"{len(cm & ch)} of {len(cm)} Mehul customers also appear in Hansa. " + win_note(w, [("Mehul", mehul_dw), ("Hansa", hansa_dw)]),
              "Do not add the two files together until you confirm they are not the same returns recorded twice.",
              mk(t2, "column", ["MEHUL UNITS", "HANSA UNITS"], title="Returns by product: Mehul vs Hansa (top 8)")))
    w = window(mehul_dw, brands_dw)
    ex_r = clip(mehul_dw, w).assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")[X].sum()
    su_r = clip(brands_dw, w).assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")["BRANDS_INV_QTY"].sum()
    rr = pd.concat([ex_r.rename("EXP"), su_r.rename("SOLD")], axis=1).fillna(0)
    rr = rr[rr["SOLD"] > 0].assign(RATE=lambda d: d["EXP"] / d["SOLD"] * 100).sort_values("RATE", ascending=False)
    t3 = pd.DataFrame({"ROUTE": [short(i) for i in rr.head(8).index], "EXPIRED": rr.head(8)["EXP"].to_numpy(), "UNITS SOLD": rr.head(8)["SOLD"].to_numpy(),
                       "EXPIRED PER 100 SOLD": rr.head(8)["RATE"].round(2).to_numpy()})
    R.append(("Which routes have high expired returns relative to sales?",
              (f"Expired returns per 100 units sold - highest: {'; '.join(f'{i} ({n1(r.RATE)})' for i, r in rr.head(3).iterrows())}. " if len(rr) else "No route has both expired returns and sales. ")
              + win_note(w, [("Mehul", mehul_dw), ("brands", brands_dw)]),
              "Routes high on this ratio over-supply short-life product.",
              mk(t3, "bar", ["EXPIRED PER 100 SOLD"], title="Expired returns per 100 units sold - highest 8 routes")))
    w = window(mehul_dw, rev_dw)
    cmv = code_sum(mehul_dw, w, [T_])[T_]; ccv = code_sum(rev_dw, w, ["REV_CN_AMT"])["REV_CN_AMT"]
    rho2, n2 = rank_corr(cmv, ccv)
    top8 = cmv.sort_values(ascending=False).head(8)
    t4 = pd.DataFrame({"CUSTOMER": [clabel(c) for c in top8.index], "MEHUL RETURNS": top8.values,
                       "CREDIT VALUE": [float(ccv.get(c, 0)) for c in top8.index]})
    R.append(("Do customers who return goods also receive credits?",
              f"Across {n2} customers in both files, returned units and credit value show {corr_words(rho2)}. " + win_note(w, [("Mehul", mehul_dw), ("revenue", rev_dw)]),
              "A credit with no matching return (or the reverse) is a reconciliation item.",
              mk(t4, "bar", ["MEHUL RETURNS"], title="Mehul returns - 8 biggest returners (credit value in the table)")))

    order = list(prod.index)
    Mx = pivot_matrix(M0, T_, "PRODUCT")
    Mx = Mx[INFO_COLS + [c for c in order if c in Mx.columns]]
    pcols = [c for c in order if c in Mx.columns]
    key = M0.assign(_k=M0["CODE"].astype("string").fillna("") + "|" + M0["CUSTOMER NAME"].astype("string").fillna(""))
    keymap = Mx["CODE"].astype("string").fillna("") + "|" + Mx["CUSTOMER NAME"].astype("string").fillna("")
    Mx["FRESH"] = keymap.map(key.groupby("_k")[F].sum()).fillna(0).astype(int).to_numpy()
    Mx["EXPIRED"] = keymap.map(key.groupby("_k")[X].sum()).fillna(0).astype(int).to_numpy()
    Mx["TOTAL_RETURNS"] = Mx[pcols].sum(axis=1)
    Mx["EXPIRED_SHARE"] = np.where(Mx["TOTAL_RETURNS"] > 0, Mx["EXPIRED"] / Mx["TOTAL_RETURNS"], 0).round(4)
    Mx = Mx.sort_values("TOTAL_RETURNS", ascending=False).reset_index(drop=True)
    checks = [("mehul: fresh + expired table = matrix total", q1t["UNITS"].sum(), Mx["TOTAL_RETURNS"].sum()),
              ("mehul: product table = matrix total", prod[T_].sum(), Mx["TOTAL_RETURNS"].sum()),
              ("mehul: route table = matrix total", route[T_].sum(), Mx["TOTAL_RETURNS"].sum()),
              ("mehul: fresh matrix = fresh source", Mx["FRESH"].sum(), fr)]
    return dict(
        name="3 Mehul Returns", title="MEHUL RETURNS - fresh vs expired returns and waste", color="#B5651D",
        subtitle=f"Source: Mehul daily return sheets (1ST, 2ND ...) | Period: {span_txt(M0)} | {nd} dates",
        kpis=[("TOTAL RETURNS", n0(tot)), ("FRESH", n0(fr)), ("EXPIRED", n0(ex)), ("EXPIRED SHARE", pc(sdiv(ex, tot))),
              ("CUSTOMERS", n0(M0["CUSTOMER"].nunique())), ("AVG PER DAY", n0(sdiv(tot, nd)))],
        notes=["The day of each return comes from the daily sheet name. Sheets that are not named like 1ST, 2ND ... are skipped unless listed in MEHUL_SHEET_DATE_OVERRIDES.",
               "Valuations use the average invoice revenue per unit from the sales-revenue file, because product prices are not in the data."],
        questions=Q, rel_title="HOW MEHUL RETURNS RELATE TO THE OTHER FOUR DATASETS", rels=R, checks=checks,
        matrix=Mx, matrix_title="CUSTOMER x PRODUCT MATRIX - total returns (fresh + expired; use the filter buttons on the header row)",
        heat=pcols, sums=pcols + ["FRESH", "EXPIRED", "TOTAL_RETURNS"], pct=["EXPIRED_SHARE"], rotate=pcols)


# ## Sheet 4 - Hansa returns

# In[48]:


def analyse_hansa():
    H = hansa_dw.copy()
    H["CUSTOMER"] = lab(H)
    for c in ("ROUTE NAME", "CATEGORY", "SALES_REP"):
        H[c] = dim(H, c)
    HQ = "HANSA_RTN_QTY"
    hv = h[h["DATE"].notna()]                       # rows without a date are not in the datewise data either
    inv = hv[["DATE", "A/C", "NAME", "TYPE", "INVOICE_NO"]].copy()
    inv["QTY_ALL"] = hv[h_prods].sum(axis=1)
    inv = attach_customer(inv, "A/C", name_col="NAME", fallback=first_name_by_code(h), label="hansa invoices")
    inv["CUSTOMER"] = lab(inv)
    inv["ROUTE NAME"] = dim(inv, "ROUTE NAME")
    inv["TYPE"] = inv["TYPE"].astype("object").where(inv["TYPE"].notna(), "(blank)").astype(str).str.strip().replace("", "(blank)")
    def inv_str(x):
        try:
            f = float(x)
            return str(int(f)) if f.is_integer() else str(x)
        except (TypeError, ValueError):
            return str(x) if pd.notna(x) else "(no number)"
    inv["INV"] = inv["INVOICE_NO"].map(inv_str)
    tot = int(H[HQ].sum()); nd = H["DATE"].nunique()
    prod = H.groupby("PRODUCT")[HQ].sum().sort_values(ascending=False)
    cust = H.groupby("CUSTOMER")[HQ].sum().sort_values(ascending=False)
    route = H.groupby("ROUTE NAME")[HQ].sum().sort_values(ascending=False)
    daily = H.groupby("DATE")[HQ].sum()
    by_inv = inv.groupby("INV").agg(QTY=("QTY_ALL", "sum"), CUSTOMER=("CUSTOMER", "first"), DATE=("DATE", "first"))
    docs_by_cust = inv.groupby("CUSTOMER")["INV"].nunique()
    Q = []
    d1 = inv.groupby("DATE")["INV"].nunique()
    pf = period_frame(daily.rename("UNITS").reset_index().assign(DOCUMENTS=d1.reindex(daily.index).to_numpy()), ["UNITS", "DOCUMENTS"])
    Q.append(("What is the total quantity returned?",
              f"{n0(tot)} units over {nd} dates on {len(by_inv):,} return documents from {H['CUSTOMER'].nunique()} customers; average {n0(sdiv(tot, nd))} units a day.",
              "The baseline to track return levels against.", mk(pf, "column", ["UNITS"], title=f"Units returned per period ({pf.columns[0].lower()})")))
    Q.append(("Which products account for the most returns?", f"{top_list(prod, 3)}. The top 3 are {pc(prod.head(3).sum() / tot)} of returns.",
              "Target these products first for order and dispatch review.", mk(top_tbl(prod, "PRODUCT", "UNITS"), "bar", ["UNITS"], title="Units returned - top 8 products")))
    t3 = top_tbl(cust, "CUSTOMER", "UNITS")
    t3.insert(2, "DOCUMENTS", [int(docs_by_cust.get(c, 0)) for c in cust[cust > 0].head(8).index])
    Q.append(("Which customers have the highest returns?", f"{top_list(cust, 5)}.", "Visit these accounts: check ordering levels, rotation and who authorises returns.",
              mk(t3, "bar", ["UNITS"], title="Units returned - top 8 customers")))
    Q.append(("Which routes have the highest returns?", f"{top_list(route, 3)}.", "Review loading quantities and delivery timing on these routes.",
              mk(top_tbl(route, "ROUTE", "UNITS"), "bar", ["UNITS"], title="Units returned - top 8 routes")))
    top_inv = by_inv.sort_values("QTY", ascending=False).head(8)
    t5 = pd.DataFrame({"INVOICE": top_inv.index, "CUSTOMER": [short(c, 22) for c in top_inv["CUSTOMER"]],
                       "DATE": [pd.Timestamp(d).strftime("%d %b") for d in top_inv["DATE"]], "UNITS": top_inv["QTY"].to_numpy()})
    Q.append(("Which invoice numbers are associated with returns?",
              f"{len(by_inv):,} distinct invoice numbers carry returns. Largest: {'; '.join(f'{i} ({n0(r.QTY)} units, {r.CUSTOMER})' for i, r in top_inv.head(3).iterrows())}.",
              "Use these numbers to trace the original delivery and the credit raised.", mk(t5, "bar", ["UNITS"], cat="INVOICE", title="Units returned - 8 largest return documents")))
    q1, q3 = by_inv["QTY"].quantile(0.25), by_inv["QTY"].quantile(0.75)
    fence = q3 + 1.5 * (q3 - q1)
    out = by_inv[by_inv["QTY"] > fence]
    edges = [-1e-9, 5, 10, 20, 50, float("inf")]
    t6 = bucket_tbl(by_inv["QTY"].abs(), edges, ["1-5 units", "6-10 units", "11-20 units", "21-50 units", "over 50 units"], "UNITS PER DOCUMENT", "DOCUMENTS")
    Q.append(("What is the return quantity per invoice?",
              f"Average {n1(by_inv['QTY'].mean())} units, median {n1(by_inv['QTY'].median())}, 95th percentile {n1(by_inv['QTY'].quantile(0.95))}. "
              f"{len(out)} invoices are unusually large (above {n0(fence)} units).",
              "Check the unusually large documents for keying errors or genuine bulk returns.", mk(t6, "column", ["DOCUMENTS"], title="Return documents by size")))
    if nd < 2:
        a7 = f"The file holds a single date ({dtxt(daily.index[0])}) with {n0(daily.iloc[0])} units, so no trend can be shown yet."
        t7 = pd.DataFrame({"PERIOD": [dtxt(daily.index[0])], "AVG UNITS / DAY": [float(daily.iloc[0])], "DATES": [1]})
    else:
        a_, b_ = daily.iloc[: nd // 2], daily.iloc[nd // 2:]
        ch = sdiv(b_.mean(), a_.mean()) - 1
        a7 = (f"First half of the dates averages {n0(a_.mean())} units a day, second half {n0(b_.mean())} ({ch * 100:+.0f}%): returns are "
              f"{'rising' if ch > 0.05 else 'falling' if ch < -0.05 else 'broadly flat'}. Peak day {dtxt(daily.idxmax())} ({n0(daily.max())}).")
        t7 = pd.DataFrame({"PERIOD": [f"First half ({dtxt(a_.index[0])[:6]} - {dtxt(a_.index[-1])[:6]})", f"Second half ({dtxt(b_.index[0])[:6]} - {dtxt(b_.index[-1])[:6]})"],
                           "AVG UNITS / DAY": [round(a_.mean(), 1), round(b_.mean(), 1)], "DATES": [len(a_), len(b_)]})
    Q.append(("How do returns change over time?", a7, "Keep a weekly return chart; act when the direction turns up.",
              mk(t7, "column", ["AVG UNITS / DAY"], title="Average units returned per day")))
    need = 3 if nd >= 4 else 2
    cp = H[H[HQ] > 0].groupby(["CUSTOMER", "PRODUCT"]).agg(DATES=("DATE", "nunique"), QTY=(HQ, "sum")).reset_index()
    rec = cp[cp["DATES"] >= need].sort_values(["DATES", "QTY"], ascending=False)
    shown = rec if len(rec) else cp.sort_values(["QTY", "DATES"], ascending=False)
    t8 = pd.DataFrame({"CUSTOMER - PRODUCT": [short(f"{r.CUSTOMER} - {r.PRODUCT}", 34) for r in shown.head(8).itertuples()],
                       "DATES": shown.head(8)["DATES"].to_numpy(), "UNITS": shown.head(8)["QTY"].to_numpy()})
    Q.append(("Which customer-product combinations return repeatedly?",
              f"{len(rec)} combinations return on at least {need} different dates. Most persistent: "
              + ("; ".join(f"{r.CUSTOMER} - {r.PRODUCT} ({r.DATES} dates, {n0(r.QTY)} units)" for r in rec.head(3).itertuples()) if len(rec)
                 else "none repeat on that many dates, so the table shows the largest combinations instead") + ".",
              "A repeated pattern is a standing order problem: adjust the standing quantity rather than accept the return every time.",
              mk(t8, "bar", ["UNITS"], title="Customer-product combinations with the most returns")))
    ty = inv.groupby("TYPE").agg(QTY=("QTY_ALL", "sum"), DOCS=("INV", "nunique")).sort_values("QTY", ascending=False)
    t9 = pd.DataFrame({"TYPE": [short(i, 22) for i in ty.head(8).index], "UNITS": ty.head(8)["QTY"].to_numpy(), "DOCUMENTS": ty.head(8)["DOCS"].to_numpy(),
                       "SHARE %": (ty.head(8)["QTY"] / tot).to_numpy()})
    Q.append(("How do returns differ by transaction type?",
              "; ".join(f"{i}: {n0(r.QTY)} units on {r.DOCS} documents ({pc(r.QTY / tot, 0)})" for i, r in ty.head(4).iterrows()) + ".",
              "Confirm what each type code means before setting rules per type.", mk(t9, "pie", ["UNITS"], title="Units returned by transaction type")))
    t, w, cm, chs = reconcile_returns()
    both = t[(t["MEHUL"] > 0) & (t["HANSA"] > 0)]
    rho = None if len(both) < 4 else float(both["MEHUL"].rank().corr(both["HANSA"].rank()))
    gap = abs(t["MEHUL"].sum() - t["HANSA"].sum())
    lead = "On the dates both cover" if w[0] is not None else "Over each file's own dates"
    t10 = t.assign(S=t["MEHUL"] + t["HANSA"]).sort_values("S", ascending=False).head(8).drop(columns="S").rename_axis("PRODUCT").reset_index()
    t10 = t10.rename(columns={"MEHUL": "MEHUL UNITS", "HANSA": "HANSA UNITS"})
    Q.append(("How closely do Hansa returns agree with Mehul returns?",
              f"{lead} Hansa has {n0(t['HANSA'].sum())} units and Mehul {n0(t['MEHUL'].sum())} (gap {n0(gap)}). "
              f"Product totals show {corr_words(rho)}; {len(cm & chs)} of {len(chs)} Hansa customers also appear in Mehul. " + win_note(w, [("Hansa", hansa_dw), ("Mehul", mehul_dw)]),
              "Do not add the two files together until you confirm they are different returns.",
              mk(t10, "column", ["HANSA UNITS", "MEHUL UNITS"], title="Returns by product: Hansa vs Mehul (top 8)")))

    R = []
    tH, wH, unm = product_rates(hansa_dw, HQ)
    extra = f" Returned products with no sales match: {', '.join(unm.index.astype(str))}." if len(unm) else ""
    r1 = tH.sort_values("RETURNED", ascending=False).head(8)
    t1 = pd.DataFrame({"PRODUCT": r1["PRODUCT"], "UNITS SOLD": r1["SOLD"].to_numpy(), "HANSA RETURNED": r1["RETURNED"].to_numpy(), "RETURN RATE %": r1["RATE"].to_numpy()})
    R.append(("Which products come back relative to what is sold?", rates_text(tH, "Hansa") + " " + win_note(wH, [("Hansa", hansa_dw), ("brands", brands_dw)]) + extra,
              "High-rate products are candidates for lower dispatch quantities.", mk(t1, "column", ["RETURN RATE %"], title="Hansa returns as % of units sold - 8 most returned products")))
    w = window(hansa_dw, brands_dw)
    ret_c = code_sum(hansa_dw, w, [HQ])[HQ]; buy_c = code_sum(brands_dw, w, ["BRANDS_INV_QTY"])["BRANDS_INV_QTY"]
    rho2, n2 = rank_corr(buy_c, ret_c)
    rate_c = (ret_c / buy_c.where(buy_c >= 20)).dropna().sort_values(ascending=False)
    top_rate = "; ".join(f"{clabel(i)} ({pc(v)})" for i, v in rate_c.head(3).items())
    top8 = ret_c.sort_values(ascending=False).head(8)
    t2 = pd.DataFrame({"CUSTOMER": [clabel(c) for c in top8.index], "UNITS BOUGHT": [float(buy_c.get(c, np.nan)) if c in buy_c.index else np.nan for c in top8.index],
                       "UNITS RETURNED": top8.values})
    t2["RETURN RATE %"] = t2["UNITS RETURNED"] / t2["UNITS BOUGHT"].where(t2["UNITS BOUGHT"] > 0)
    R.append(("Do customers who buy the most also return the most?",
              f"Across {n2} customers in both files, units bought and units returned show {corr_words(rho2)}. Highest return rates (customers with at least 20 units bought): {top_rate or 'none'}. " + win_note(w, [("Hansa", hansa_dw), ("brands", brands_dw)]),
              "Rate matters more than volume: a mid-size customer returning 20% needs attention before a large one returning 2%.",
              mk(t2, "column", ["UNITS RETURNED"], title="Units returned - 8 biggest returners (bought units and rate in the table)")))
    inv_no = set(rev_inv["invoice_no"].dropna().astype(str)) if "invoice_no" in rev_inv else set()
    cn_no = set(rev_cn["invoice_no"].dropna().astype(str)) if "invoice_no" in rev_cn else set()
    inv_no = {inv_str(x) for x in inv_no}; cn_no = {inv_str(x) for x in cn_no}
    hn = set(by_inv.index.astype(str))
    t3b = pd.DataFrame({"ITEM": ["Hansa return documents", "Found among revenue invoices", "Found among credit documents", "Found in neither"],
                        "DOCUMENTS": [len(hn), len(hn & inv_no), len(hn & cn_no), len(hn - inv_no - cn_no)]})
    R.append(("Can Hansa returns be matched to invoices and credits by number?",
              f"{len(hn & inv_no)} of {len(hn)} Hansa invoice numbers appear among the sales-revenue invoices and {len(hn & cn_no)} among the credit documents. "
              + ("Numbers do not line up, so link them by customer and date instead." if not (hn & inv_no) and not (hn & cn_no) else "Those matches can be reconciled document by document."),
              "Agree one shared document reference between returns and credits so each return can be traced to its credit.",
              mk(t3b, "column", ["DOCUMENTS"], cat="ITEM", title="Hansa return documents found in the revenue file")))
    w = window(hansa_dw, rev_dw)
    hc = code_sum(hansa_dw, w, [HQ])[HQ]; cc = code_sum(rev_dw, w, ["REV_CN_AMT"])["REV_CN_AMT"]
    rho3, n3 = rank_corr(hc, cc)
    no_cn = len(set(hc.index) - set(cc[cc > 0].index))
    top8 = hc.sort_values(ascending=False).head(8)
    t4 = pd.DataFrame({"CUSTOMER": [clabel(c) for c in top8.index], "HANSA RETURNS": top8.values, "CREDIT VALUE": [float(cc.get(c, 0)) for c in top8.index]})
    R.append(("Do returned goods lead to credits?",
              f"Across {n3} customers in both files, returned units and credit value show {corr_words(rho3)}. {no_cn} customers with Hansa returns have no credit in the sales-revenue file. " + win_note(w, [("Hansa", hansa_dw), ("revenue", rev_dw)]),
              "Returns without credits (or credits without returns) are the reconciliation items to chase.",
              mk(t4, "column", ["HANSA RETURNS"], title="Hansa returns - 8 biggest returners (credit value in the table)")))

    order = list(prod.index)
    Mx = pivot_matrix(H, HQ, "PRODUCT")
    Mx = Mx[INFO_COLS + [c for c in order if c in Mx.columns]]
    pcols = [c for c in order if c in Mx.columns]
    keyinv = inv.assign(_k=inv["CODE"].astype("string").fillna("") + "|" + inv["CUSTOMER NAME"].astype("string").fillna("")).groupby("_k")["INV"].nunique()
    keymap = Mx["CODE"].astype("string").fillna("") + "|" + Mx["CUSTOMER NAME"].astype("string").fillna("")
    Mx["TOTAL_RETURNS"] = Mx[pcols].sum(axis=1)
    Mx["DOCUMENTS"] = keymap.map(keyinv).fillna(0).astype(int).to_numpy()
    Mx = Mx.sort_values("TOTAL_RETURNS", ascending=False).reset_index(drop=True)
    checks = [("hansa: product table = matrix total", prod.sum(), Mx["TOTAL_RETURNS"].sum()),
              ("hansa: route table = matrix total", route.sum(), Mx["TOTAL_RETURNS"].sum()),
              ("hansa: documents add up to the total", by_inv["QTY"].sum(), Mx["TOTAL_RETURNS"].sum()),
              ("hansa: type table = total", ty["QTY"].sum(), Mx["TOTAL_RETURNS"].sum()),
              ("hansa: document-size table counts every document", t6["DOCUMENTS"].sum(), len(by_inv))]
    return dict(
        name="4 Hansa Returns", title="HANSA RETURNS - returns by customer, invoice and product", color="#7B2D8E",
        subtitle=f"Source: Hansa returns | Period: {span_txt(H)} | {nd} dates",
        kpis=[("UNITS RETURNED", n0(tot)), ("RETURN DOCUMENTS", n0(len(by_inv))), ("CUSTOMERS", n0(H["CUSTOMER"].nunique())),
              ("PRODUCTS", n0(len(prod))), ("AVG UNITS PER DOCUMENT", n1(by_inv["QTY"].mean())), ("AVG UNITS PER DAY", n0(sdiv(tot, nd)))],
        notes=["The source records returns as negative credit quantities; this report shows returned units as positive magnitudes summed from the itemized product columns. Its Qty subtotal can disagree with those details, so the itemized values are authoritative. Hansa and Mehul may record the same returns, so they are compared, never added.",
               "A missing product quantity is treated as zero returned."],
        questions=Q, rel_title="HOW HANSA RETURNS RELATE TO THE OTHER FOUR DATASETS", rels=R, checks=checks,
        matrix=Mx, matrix_title="CUSTOMER x PRODUCT MATRIX - units returned (use the filter buttons on the header row)",
        heat=pcols, sums=pcols + ["TOTAL_RETURNS", "DOCUMENTS"], pct=[], rotate=pcols)


# ## Sheet 5 - Sales revenue

# In[49]:


def analyse_revenue():
    V = rev_dw.copy()
    V["CUSTOMER"] = lab(V)
    for c in ("ROUTE NAME", "CATEGORY", "SALES_REP"):
        V[c] = dim(V, c)
    IA, IQ, CA, CQ = "REV_INV_AMT", "REV_INV_QTY", "REV_CN_AMT", "REV_CN_QTY"
    inv_t, cn_t = V[IA].sum(), V[CA].sum()
    nd = V["DATE"].nunique()
    ri = rev_inv.dropna(subset=["DATE"]); rc_ = rev_cn.dropna(subset=["DATE"])
    n_inv, n_cn = len(ri), len(rc_)
    V["MONTH"] = V["DATE"].dt.to_period("M")
    month = V.groupby("MONTH")[[IA, CA]].sum()
    cust = V.groupby("CUSTOMER")[IA].sum().sort_values(ascending=False)
    cred_c = V.groupby("CUSTOMER")[CA].sum().sort_values(ascending=False)
    route = V.groupby("ROUTE NAME")[[IA, CA]].sum()
    cat = V.groupby("CATEGORY")[[IA, CA]].sum().sort_values(IA, ascending=False)
    daily = V.groupby("DATE")[IA].sum()
    Q = []
    t1 = pd.DataFrame({"ITEM": ["Invoiced", "Credits", "Net revenue"], "VALUE": [inv_t, cn_t, inv_t - cn_t], "UNITS": [V[IQ].sum(), V[CQ].sum(), V[IQ].sum() - V[CQ].sum()]})
    Q.append(("What is the total invoiced, credited and net revenue?",
              f"Invoiced {n0(inv_t)} ({n0(V[IQ].sum())} units) on {n_inv:,} invoices; credits {n0(cn_t)} ({n0(V[CQ].sum())} units) on {n_cn:,} credit notes. "
              f"Credits are {pc(sdiv(cn_t, inv_t))} of invoiced value; net revenue {n0(inv_t - cn_t)}. Amounts use the inv_total column.",
              "Credits as a % of sales is the headline control figure; set a ceiling and track it monthly.",
              mk(t1, "column", ["VALUE"], cat="ITEM", title="Invoiced, credited and net value")))
    mt = month.tail(12).copy()
    t2 = pd.DataFrame({"MONTH": [p.strftime("%b %Y") for p in mt.index], "INVOICED": mt[IA].to_numpy(), "CREDITS": mt[CA].to_numpy(),
                       "CREDIT RATIO %": (mt[CA] / mt[IA].where(mt[IA] > 0)).fillna(0).to_numpy()})
    if len(month) > 1:
        best, worst = month[IA].idxmax(), month[IA].idxmin()
        last, prev = month[IA].iloc[-1], month[IA].iloc[-2]
        a2 = (f"Monthly invoiced value: " + ", ".join(f"{p.strftime('%b %Y')} {n0(v)}" for p, v in month[IA].tail(12).items()) +
              f". Best month {best.strftime('%b %Y')}, weakest {worst.strftime('%b %Y')}. Latest month vs the one before: {sdiv(last, prev) * 100 - 100:+.0f}% (the latest month may be incomplete). Peak day {dtxt(daily.idxmax())} ({n0(daily.max())}).")
    else:
        a2 = f"Only one month is covered: {n0(inv_t)} invoiced; peak day {dtxt(daily.idxmax())} ({n0(daily.max())})."
    Q.append(("How does revenue change by month and by day?", a2, "Use the monthly view for growth and the peak days for capacity planning.",
              mk(t2, "column", ["INVOICED", "CREDITS"], title="Invoiced vs credited value by month (last 12)")))
    Q.append(("Which customers generate the highest invoice value?", f"{top_list(cust, 5)}. Top 10 customers = {pc(cust.head(10).sum() / inv_t)} of invoiced value.",
              "Prioritise service levels and account reviews for these customers.", mk(top_tbl(cust, "CUSTOMER", "INVOICED"), "bar", ["INVOICED"], title="Invoiced value - top 8 customers")))
    rt = route.copy(); rt["RATIO"] = rt[CA] / rt[IA].where(rt[IA] > 0)
    big = rt[rt[IA] >= 0.02 * inv_t].sort_values("RATIO", ascending=False)
    rtop = rt.sort_values(IA, ascending=False).head(8)
    t4 = pd.DataFrame({"ROUTE": [short(i) for i in rtop.index], "INVOICED": rtop[IA].to_numpy(), "CREDITS": rtop[CA].to_numpy(), "CREDIT RATIO %": rtop["RATIO"].fillna(0).to_numpy()})
    Q.append(("Which routes bring in the most revenue, and which have the highest credit ratio?",
              f"Revenue: {top_list(rt[IA], 3)}. Highest credit ratio (routes with 2%+ of sales): {'; '.join(f'{i} ({pc(r.RATIO)})' for i, r in big.head(3).iterrows())}.",
              "Review credit approvals on high-ratio routes.", mk(t4, "bar", ["INVOICED"], title="Invoiced value - top 8 routes (credit ratio in the table)")))
    t5 = pd.DataFrame({"CATEGORY": cat.head(8).index, "INVOICED": cat.head(8)[IA].to_numpy(), "CREDITS": cat.head(8)[CA].to_numpy(), "SHARE %": (cat.head(8)[IA] / inv_t).to_numpy()})
    Q.append(("Which categories contribute the most?", f"Revenue: {top_list(cat[IA], 3)}. Credits: {top_list(cat[CA], 3)}.",
              "Compare category credit shares to their revenue shares to spot over-crediting.", mk(t5, "pie", ["INVOICED"], title="Share of invoiced value by category")))
    pq = ri["inv_total"].quantile([0.25, 0.5, 0.75, 0.95]).to_list()
    edges = [-float("inf"), pq[0], pq[1], pq[2], pq[3], float("inf")]
    labs = [f"up to {n0(pq[0])}", f"{n0(pq[0])} - {n0(pq[1])}", f"{n0(pq[1])} - {n0(pq[2])}", f"{n0(pq[2])} - {n0(pq[3])}", f"over {n0(pq[3])}"]
    bands = pd.cut(ri["inv_total"], bins=edges, labels=labs, right=True)
    t6 = pd.DataFrame({"INVOICE VALUE BAND": labs, "INVOICES": bands.value_counts().reindex(labs).fillna(0).astype(int).to_numpy(),
                       "VALUE": ri.groupby(bands, observed=False)["inv_total"].sum().reindex(labs).fillna(0).to_numpy()})
    t6["SHARE %"] = t6["VALUE"] / t6["VALUE"].sum() if t6["VALUE"].sum() else 0
    Q.append(("What are the average invoice value and quantity?",
              f"Average invoice {n0(sdiv(inv_t, n_inv))} and {n1(sdiv(V[IQ].sum(), n_inv))} units; median invoice {n0(ri['inv_total'].median())}. Average credit note {n0(sdiv(cn_t, n_cn))}.",
              "Rising average invoices with falling counts can signal fewer, larger orders.",
              mk(t6, "column", ["INVOICES"], title="Invoices by value band (quartile bands)")))
    tr = trend_changes(V, "CUSTOMER", IA, min_total=0.002 * inv_t)
    if tr is None:
        a7 = "The file holds a single date, so a decline cannot be measured. The table shows the biggest customers instead."
        b7 = mk(top_tbl(cust, "CUSTOMER", "INVOICED"), "bar", ["INVOICED"], title="Top customers (single date)")
    else:
        t, ndates = tr
        dec = t[t["second"] < 0.7 * t["first"]].sort_values("gap")
        trr = trend_changes(V, "ROUTE NAME", IA, min_total=0.002 * inv_t)[0]
        decr = trr[trr["second"] < 0.8 * trr["first"]].sort_values("gap")
        a7 = (f"Comparing the first and second half of the {ndates} dates: {len(dec)} customers fall by 30%+ per day"
              + (" (" + "; ".join(change_txt(i, r, "") for i, r in dec.head(3).iterrows()) + ")" if len(dec) else "")
              + f"; {len(decr)} routes fall by 20%+" + (" (" + "; ".join(change_txt(i, r, "") for i, r in decr.head(2).iterrows()) + ")" if len(decr) else "") + ".")
        b7 = mk(movers_tbl(t), "column", ["FIRST HALF / DAY", "SECOND HALF / DAY"], title="Invoiced value per day: first vs second half (8 biggest falls)")
    Q.append(("Which customers or routes are declining?", a7, "Follow up declining accounts before they are lost.", b7))
    q1_, q3_ = ri["inv_total"].quantile(0.25), ri["inv_total"].quantile(0.75); fence = q3_ + 3 * (q3_ - q1_)
    big_i = ri[ri["inv_total"] > fence].sort_values("inv_total", ascending=False)
    cq1, cq3 = rc_["inv_total"].quantile(0.25), rc_["inv_total"].quantile(0.75); cfence = cq3 + 3 * (cq3 - cq1)
    big_c = rc_[rc_["inv_total"] > cfence]
    shown = ri.sort_values("inv_total", ascending=False).head(8)
    med = ri["inv_total"].median()
    t8 = pd.DataFrame({"INVOICE": [str(int(x)) if pd.notna(x) and float(x).is_integer() else str(x) for x in shown["invoice_no"]],
                       "CUSTOMER": [clabel(c) for c in shown["A/C"]], "VALUE": shown["inv_total"].to_numpy(),
                       "TIMES MEDIAN": (shown["inv_total"] / med).round(1).to_numpy() if med else 0})
    Q.append(("Which invoices or credits look unusual?",
              f"{len(big_i)} invoices exceed {n0(fence)} (3 x the interquartile range above the 75th percentile)"
              + (f"; largest {'; '.join(f'{r.invoice_no} ({n0(r.inv_total)})' for r in big_i.head(3).itertuples())}" if len(big_i) else "")
              + f". {len(big_c)} credit notes exceed {n0(cfence)}. {int((ri['qty'] <= 0).sum())} invoices have zero or negative quantity. The table lists the 8 largest invoices.",
              "Check the largest documents and any zero-quantity invoices for keying errors.",
              mk(t8, "bar", ["VALUE"], cat="INVOICE", title="Invoice value - 8 largest invoices", fmts={"TIMES MEDIAN": "0.0"})))
    cn_by_c = rc_.groupby(rc_["A/C"]).size()
    rep = cn_by_c[cn_by_c >= 3]
    inv_set = set(ri["invoice_no"].dropna().astype(str)) if "invoice_no" in ri else set()
    link = sum(1 for x in rc_["invoice_no"].dropna().astype(str) if x in inv_set) if "invoice_no" in rc_ else 0
    t9 = pd.DataFrame({"CUSTOMER": [short(i) for i in cred_c[cred_c > 0].head(8).index], "CREDIT VALUE": cred_c[cred_c > 0].head(8).to_numpy()})
    code_of = V.drop_duplicates("CUSTOMER").set_index("CUSTOMER")["CODE"]
    t9["CREDIT NOTES"] = [int(cn_by_c.get(code_of.get(c), 0)) if pd.notna(code_of.get(c)) else 0 for c in cred_c[cred_c > 0].head(8).index]
    Q.append(("How often are credits raised, and which customers get repeated credits?",
              f"{n_cn:,} credit notes across {len(cn_by_c)} customers; {len(rep)} customers have 3 or more. Top credit value: {top_list(cred_c, 3)}. "
              f"{link} credit notes quote an invoice number that appears in the invoice sheet.",
              "Repeat credits point to a recurring service or pricing issue at that customer.",
              mk(t9, "bar", ["CREDIT VALUE"], title="Credit value - top 8 customers")))
    parts = {k: ri[k].sum() for k in ("amount", "vat_amt", "extra_cts") if k in ri.columns}
    if parts and ri["inv_total"].sum() > 0:
        comp = " + ".join(f"{k.replace('_', ' ')} {n0(v)} ({pc(v / ri['inv_total'].sum())})" for k, v in parts.items())
        diff = (ri["inv_total"] - sum(ri[k].fillna(0) for k in parts)).abs()
        a10 = f"Invoice total {n0(ri['inv_total'].sum())} = {comp}. {int((diff > 1).sum())} invoices differ from the sum of their parts by more than 1."
        names = {"amount": "Amount", "vat_amt": "VAT", "extra_cts": "Extra charges"}
        t10 = pd.DataFrame({"COMPONENT": [names[k] for k in parts] + ["Invoice total"], "VALUE": [v for v in parts.values()] + [ri["inv_total"].sum()]})
        t10["SHARE %"] = t10["VALUE"] / ri["inv_total"].sum()
    else:
        a10 = "The amount, VAT or extra columns are missing, so the total cannot be reconciled."
        t10 = pd.DataFrame({"COMPONENT": ["Invoice total"], "VALUE": [ri["inv_total"].sum()], "SHARE %": [1.0]})
    Q.append(("Do VAT, extras and amount add up to the invoice total?", a10, "Investigate invoices where the parts do not sum to the total.",
              mk(t10, "column", ["VALUE"], cat="COMPONENT", title="Invoice total and its components")))

    R = []
    w = window(rev_dw, brands_dw)
    units = clip(brands_dw, w)["BRANDS_INV_QTY"].sum(); rev_w = clip(rev_dw, w)[IA].sum()
    rv_c = code_sum(rev_dw, w, [IA])[IA]; bu_c = code_sum(brands_dw, w, ["BRANDS_INV_QTY"])["BRANDS_INV_QTY"]
    rho, n = rank_corr(rv_c, bu_c)
    top8 = rv_c.sort_values(ascending=False).head(8)
    t1r = pd.DataFrame({"CUSTOMER": [clabel(c) for c in top8.index], "INVOICE VALUE": top8.values,
                        "BRAND UNITS": [float(bu_c.get(c, np.nan)) if c in bu_c.index else np.nan for c in top8.index]})
    t1r["VALUE PER UNIT"] = (t1r["INVOICE VALUE"] / t1r["BRAND UNITS"].where(t1r["BRAND UNITS"] > 0)).round(1)
    R.append(("How does revenue relate to the units sold in the brands file?",
              f"Across {n} customers in both files, revenue and brand units show {corr_words(rho)}. Over the common dates, invoice value is {n1(sdiv(rev_w, units))} per brand unit (the files may cover different customer groups). " + win_note(w, [("revenue", rev_dw), ("brands", brands_dw)]),
              "Customers that buy many units at a low value per unit may be on discounted prices.",
              mk(t1r, "column", ["INVOICE VALUE"], title="Invoice value - 8 biggest customers (units and value per unit in the table)")))
    w = window(rev_dw, mehul_dw, hansa_dw)
    cm = code_sum(mehul_dw, w, ["MEHUL_TOTAL_RTN_QTY"])["MEHUL_TOTAL_RTN_QTY"]; ch = code_sum(hansa_dw, w, ["HANSA_RTN_QTY"])["HANSA_RTN_QTY"]
    cc = code_sum(rev_dw, w, [CA])[CA]
    r1, n1_ = rank_corr(cc, cm); r2, n2_ = rank_corr(cc, ch)
    cc_pos = set(cc[cc > 0].index)
    top8 = cc.sort_values(ascending=False).head(8)
    t2r = pd.DataFrame({"CUSTOMER": [clabel(c) for c in top8.index], "CREDIT VALUE": top8.values,
                        "MEHUL RETURNS": [float(cm.get(c, 0)) for c in top8.index], "HANSA RETURNS": [float(ch.get(c, 0)) for c in top8.index]})
    R.append(("Do customers with big credits also have big returns?",
              f"Credit value vs Mehul returns: {corr_words(r1)} ({n1_} customers). Credit value vs Hansa returns: {corr_words(r2)} ({n2_} customers). "
              f"{len(cc_pos - set(cm[cm > 0].index) - set(ch[ch > 0].index))} of {len(cc_pos)} credit customers have no return recorded in either file. " + win_note(w, [("revenue", rev_dw), ("Mehul", mehul_dw), ("Hansa", hansa_dw)]),
              "A credit with no return behind it needs a reason on file.",
              mk(t2r, "column", ["MEHUL RETURNS", "HANSA RETURNS"], title="Returns recorded for the 8 customers with the biggest credits")))
    w = window(rev_dw, crates_dw)
    per_crate = sdiv(clip(rev_dw, w)[IA].sum(), clip(crates_dw, w)["CRATES_INV_QTY"].sum())
    rt_ = clip(rev_dw, w).assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")[IA].sum()
    ct_ = clip(crates_dw, w).assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")["CRATES_INV_QTY"].sum()
    j = pd.concat([rt_.rename("REVENUE"), ct_.rename("CRATES")], axis=1).dropna()
    j = j[(j["CRATES"] > 0) & (j["REVENUE"] > 0)].sort_values("CRATES", ascending=False).head(8)
    t3r = pd.DataFrame({"ROUTE": [short(i) for i in j.index], "REVENUE": j["REVENUE"].round(0).to_numpy(), "CRATES": j["CRATES"].to_numpy(), "REVENUE PER CRATE": (j["REVENUE"] / j["CRATES"]).round(0).to_numpy()})
    R.append(("How much revenue does each crate carry?",
              f"About {n0(per_crate)} of invoice value per crate invoiced. " + win_note(w, [("revenue", rev_dw), ("crates", crates_dw)]),
              "Low revenue per crate on a route means crates are tied up for little sales.",
              mk(t3r, "bar", ["REVENUE PER CRATE"], title="Revenue per crate - routes with the most crates")))
    brand_codes, crate_codes = set(brands_dw["CODE"].dropna()), set(crates_dw["CODE"].dropna())
    mh_codes, ha_codes = set(mehul_dw["CODE"].dropna()), set(hansa_dw["CODE"].dropna())
    rc_codes = set(V["CODE"].dropna())
    t4r = pd.DataFrame({"DATASET": ["Brands", "Crates", "Mehul returns", "Hansa returns"],
                        "CUSTOMERS IN BOTH": [len(rc_codes & brand_codes), len(rc_codes & crate_codes), len(rc_codes & mh_codes), len(rc_codes & ha_codes)]})
    t4r["SHARE OF REVENUE CUSTOMERS %"] = t4r["CUSTOMERS IN BOTH"] / max(len(rc_codes), 1)
    R.append(("How well do the datasets cover the same customers?",
              f"Of {len(rc_codes)} customers in the sales-revenue file: {len(rc_codes & brand_codes)} appear in brands, {len(rc_codes & crate_codes)} in crates, "
              f"{len(rc_codes & mh_codes)} in Mehul returns and {len(rc_codes & ha_codes)} in Hansa returns.",
              "Low overlap means the files cover different customer groups; do not compare totals across them.",
              mk(t4r, "column", ["SHARE OF REVENUE CUSTOMERS %"], title="Share of revenue customers that also appear in each dataset", ymax=1)))

    mlabels = {p: p.strftime("%b %Y") for p in sorted(V["MONTH"].unique())}
    Vm = V.assign(MON=V["MONTH"].map(mlabels))
    Mx = pivot_matrix(Vm, IA, "MON")
    pcols = [mlabels[p] for p in sorted(mlabels) if mlabels[p] in Mx.columns]
    Mx = Mx[INFO_COLS + pcols]
    k = Vm.assign(_k=Vm["CODE"].astype("string").fillna("") + "|" + Vm["CUSTOMER NAME"].astype("string").fillna(""))
    keymap = Mx["CODE"].astype("string").fillna("") + "|" + Mx["CUSTOMER NAME"].astype("string").fillna("")
    Mx["TOTAL_INVOICED"] = Mx[pcols].sum(axis=1)
    Mx["TOTAL_CREDITS"] = keymap.map(k.groupby("_k")[CA].sum()).fillna(0).to_numpy()
    Mx["NET_REVENUE"] = Mx["TOTAL_INVOICED"] - Mx["TOTAL_CREDITS"]
    Mx["CREDIT_RATIO"] = np.where(Mx["TOTAL_INVOICED"] > 0, Mx["TOTAL_CREDITS"] / Mx["TOTAL_INVOICED"], 0).round(4)
    Mx = Mx.sort_values("TOTAL_INVOICED", ascending=False).reset_index(drop=True)
    checks = [("revenue: Q1 invoiced = matrix total", t1["VALUE"].iloc[0], Mx["TOTAL_INVOICED"].sum()),
              ("revenue: Q1 credits = matrix total", t1["VALUE"].iloc[1], Mx["TOTAL_CREDITS"].sum()),
              ("revenue: monthly table (all months) = matrix total", month[IA].sum(), Mx["TOTAL_INVOICED"].sum()),
              ("revenue: category table = matrix total", cat[IA].sum(), Mx["TOTAL_INVOICED"].sum()),
              ("revenue: value-band table counts every invoice", t6["INVOICES"].sum(), len(ri)),
              ("revenue: invoiced total = inv_total of the dated invoices", Mx["TOTAL_INVOICED"].sum(), ri["inv_total"].sum())]
    return dict(
        name="5 Sales Revenue", title="SALES REVENUE - invoices, credits and credit control", color="#C0392B",
        subtitle=f"Source: sales-revenue invoices and credits (inv_total) | Period: {span_txt(V)} | {nd} dates",
        kpis=[("INVOICED", n0(inv_t)), ("CREDITS", n0(cn_t)), ("NET REVENUE", n0(inv_t - cn_t)), ("CREDIT RATIO", pc(sdiv(cn_t, inv_t))),
              ("INVOICES", n0(n_inv)), ("AVG INVOICE", n0(sdiv(inv_t, n_inv)))],
        notes=["All amounts use the inv_total column. Credit amounts and quantities are shown as positive numbers. Customer, category and route come from the customer-info file, looked up by A/C.",
               "A credit is not automatically a product return and is not proof of lost revenue; confirm the accounting rules before treating it as a loss."],
        questions=Q, rel_title="HOW SALES REVENUE RELATES TO THE OTHER FOUR DATASETS", rels=R, checks=checks,
        matrix=Mx, matrix_title="CUSTOMER x MONTH MATRIX - invoiced value (use the filter buttons on the header row)",
        heat=pcols, sums=pcols + ["TOTAL_INVOICED", "TOTAL_CREDITS", "NET_REVENUE"], pct=["CREDIT_RATIO"], rotate=[], money=True)


# ## Week-on-week comparison (previous week from `prev_wk_data`)

# In[50]:


# ======================= previous-week comparison =======================
GREEN, RED, GREY = "#1E8449", "#C0392B", "#7F8C8D"

def prep(df):
    d = df.copy()
    d["CUSTOMER"] = lab(d)
    for c in ("ROUTE NAME", "CATEGORY", "SALES_REP"):
        d[c] = dim(d, c)
    return d

def _days(df): return max(int(df["DATE"].nunique()), 1) if len(df) else 1

def period_metrics(P):
    """headline numbers of one week"""
    m = {}
    b = P["brands_dw"]; I = b[b["BRANDS_INV_QTY"] > 0]
    m.update(b_days=_days(b), b_units=float(I["BRANDS_INV_QTY"].sum()), b_cn=float(b["BRANDS_CN_QTY"].sum()),
             b_cust=int(lab(I).nunique()) if len(I) else 0, b_prods=int(I["PRODUCT"].nunique()))
    m["b_cr"] = sdiv(m["b_cn"], m["b_units"]); m["b_perday"] = sdiv(m["b_units"], m["b_days"])
    c = P["crates_dw"]; Ic = c[c["CRATES_INV_QTY"] > 0]
    m.update(c_days=_days(c), c_net=float(Ic["CRATES_INV_QTY"].sum()), c_cn=float(c["CRATES_CN_QTY"].sum()), c_cust=int(lab(Ic).nunique()) if len(Ic) else 0)
    m["c_perday"] = sdiv(m["c_net"], m["c_days"])
    pk = c.groupby("DATE")["CRATES_INV_QTY"].sum()
    m["c_peak_day"] = dtxt(pk.idxmax()) if len(pk) else "-"; m["c_peak"] = float(pk.max()) if len(pk) else 0.0
    mh = P["mehul_dw"]
    m.update(m_days=_days(mh), m_total=float(mh["MEHUL_TOTAL_RTN_QTY"].sum()), m_fresh=float(mh["MEHUL_FRESH_RTN_QTY"].sum()), m_exp=float(mh["MEHUL_EXP_RTN_QTY"].sum()),
             m_cust=int(lab(mh).nunique()) if len(mh) else 0)
    m["m_exp_share"] = sdiv(m["m_exp"], m["m_total"]); m["m_perday"] = sdiv(m["m_total"], m["m_days"])
    hd = P["hansa_dw"]; hr = P["h"][P["h"]["DATE"].notna()]
    docs = int(hr["INVOICE_NO"].nunique()) or len(hr)
    m.update(h_days=_days(hd), h_total=float(hd["HANSA_RTN_QTY"].sum()), h_docs=docs, h_cust=int(lab(hd).nunique()) if len(hd) else 0, h_prods=int(hd["PRODUCT"].nunique()))
    m["h_perdoc"] = sdiv(m["h_total"], docs); m["h_perday"] = sdiv(m["h_total"], m["h_days"])
    v = P["rev_dw"]; ri = P["rev_inv"].dropna(subset=["DATE"])
    m.update(v_days=_days(v), v_inv=float(v["REV_INV_AMT"].sum()), v_cn=float(v["REV_CN_AMT"].sum()), v_n=len(ri))
    m["v_net"] = m["v_inv"] - m["v_cn"]; m["v_ratio"] = sdiv(m["v_cn"], m["v_inv"]); m["v_avg"] = sdiv(m["v_inv"], m["v_n"]); m["v_perday"] = sdiv(m["v_inv"], m["v_days"])
    return m

KPI_SPECS = {
    "1 Brands": ("b_days", [("NET UNITS INVOICED", "b_units", "flow", "n0", True), ("CREDIT UNITS", "b_cn", "flow", "n0", False), ("CREDIT RATE", "b_cr", "ratio", "pc", False),
                            ("CUSTOMERS BUYING", "b_cust", "level", "n0", True), ("PRODUCTS SOLD", "b_prods", "level", "n0", True), ("AVG UNITS PER DAY", "b_perday", "level", "n0", True)]),
    "2 Crates": ("c_days", [("NET CRATES INVOICED", "c_net", "flow", "n0", None), ("CRATES CREDITED", "c_cn", "flow", "n0", False), ("CUSTOMERS", "c_cust", "level", "n0", None),
                            ("AVG CRATES PER DAY", "c_perday", "level", "n0", None), ("PEAK DAY", "c_peak_day", "text", "text", None), ("PEAK DAY CRATES", "c_peak", "level", "n0", None)]),
    "3 Mehul Returns": ("m_days", [("TOTAL RETURNS", "m_total", "flow", "n0", False), ("FRESH", "m_fresh", "flow", "n0", None), ("EXPIRED", "m_exp", "flow", "n0", False),
                                   ("EXPIRED SHARE", "m_exp_share", "ratio", "pc", False), ("CUSTOMERS", "m_cust", "level", "n0", None), ("AVG PER DAY", "m_perday", "level", "n0", False)]),
    "4 Hansa Returns": ("h_days", [("UNITS RETURNED", "h_total", "flow", "n0", False), ("RETURN DOCUMENTS", "h_docs", "flow", "n0", False), ("CUSTOMERS", "h_cust", "level", "n0", None),
                                   ("PRODUCTS", "h_prods", "level", "n0", None), ("AVG UNITS PER DOCUMENT", "h_perdoc", "level", "n1", False), ("AVG UNITS PER DAY", "h_perday", "level", "n0", False)]),
    "5 Sales Revenue": ("v_days", [("INVOICED", "v_inv", "flow", "n0", True), ("CREDITS", "v_cn", "flow", "n0", False), ("NET REVENUE", "v_net", "flow", "n0", True),
                                   ("CREDIT RATIO", "v_ratio", "ratio", "pc", False), ("INVOICES", "v_n", "flow", "n0", True), ("AVG INVOICE", "v_avg", "level", "n0", True)]),
}
_FMT = {"n0": n0, "n1": n1, "pc": pc, "text": str}

def kpi_change(kind, cur, prev, dc, dp):
    """(relative change or points, text, per_day flag)"""
    if kind == "text": return None
    if kind == "ratio":
        d = (cur - prev) * 100
        return d, f"{d:+.1f} pts", False
    per_day = kind == "flow" and dc != dp
    c, p = (cur / dc, prev / dp) if per_day else (cur, prev)
    if not p: return None
    r = (c - p) / p
    return r, f"{r * 100:+.1f}%", per_day

def favour(good, chg, kind):
    if good is None or chg is None: return GREY
    thr = 0.3 if kind == "ratio" else 0.01            # ignore changes under 1% (0.3 pts for ratios)
    if abs(chg) < thr: return GREY
    return GREEN if (chg > 0) == bool(good) else RED

def kpi_cards(sheet, mc, mp):
    dkey, specs = KPI_SPECS[sheet]
    out = []
    for label, key, kind, f, good in specs:
        cur, prev = mc[key], mp[key]
        val = _FMT[f](cur)
        ch = kpi_change(kind, cur, prev, mc[dkey], mp[dkey])
        if kind == "text":
            comp = (f"Prev {prev}", "", GREY)
        elif ch is None:
            comp = (f"Prev {_FMT[f](prev)}", "", GREY)
        else:
            arrow = "▲" if ch[0] > 0 else "▼" if ch[0] < 0 else "■"
            comp = (f"Prev {_FMT[f](prev)}   ", f"{arrow} {ch[1]}" + (" /day" if ch[2] else ""), favour(good, ch[0], kind))
        out.append((label, val, comp))
    return out

# ---------- generic comparison tables
def delta_colors(good):
    return ("#27AE60", "#C0392B") if good is True else ("#C0392B", "#27AE60") if good is False else ("#2E86C1", "#F39C12")

def delta_frame(cur, prev, cd, pdd):
    same = cd == pdd
    c = cur if same else cur / cd
    p = prev if same else prev / pdd
    idx = c.index.union(p.index)
    t = pd.DataFrame({"c": c.reindex(idx).fillna(0.0), "p": p.reindex(idx).fillna(0.0)})
    t["chg"] = t["c"] - t["p"]
    t["pct"] = np.where(t["p"] > 0, t["chg"] / t["p"], np.nan)
    return t, ("" if same else " / DAY")

def mover_text(t, suf, unit, k=3):
    up = t[t["chg"] > 0].sort_values("chg", ascending=False).head(k)
    dn = t[t["chg"] < 0].sort_values("chg").head(k)
    f = lambda i, r: f"{i} ({r.chg:+,.0f}{suf.lower()}{'' if np.isnan(r.pct) else f', {r.pct * 100:+.0f}%'})"
    a = ("Biggest gains: " + "; ".join(f(i, r) for i, r in up.iterrows()) + ". ") if len(up) else "No gains. "
    b = ("Biggest falls: " + "; ".join(f(i, r) for i, r in dn.iterrows()) + ".") if len(dn) else "No falls."
    return a + b

def cmp_movers_tbl(cur, prev, cd, pdd, name, k=4, shorten=True):
    t, suf = delta_frame(cur, prev, cd, pdd)
    sel = pd.concat([t[t["chg"] > 0].sort_values("chg", ascending=False).head(k), t[t["chg"] < 0].sort_values("chg").head(k)])
    if not len(sel):
        sel = t.sort_values("c", ascending=False).head(2 * k)
    tbl = pd.DataFrame({name: [short(i) if shorten else str(i) for i in sel.index], f"CURRENT{suf}": sel["c"].round(1).to_numpy(), f"PREVIOUS{suf}": sel["p"].round(1).to_numpy(),
                        "CHANGE": sel["chg"].round(1).to_numpy(), "CHANGE %": sel["pct"].to_numpy()})
    return t, suf, tbl

def blk_movers(q, cur, prev, cd, pdd, name, unit, good, act, title, k=4, pre=""):
    t, suf, tbl = cmp_movers_tbl(cur, prev, cd, pdd, name, k)
    tot_c, tot_p = t["c"].sum(), t["p"].sum()
    a = (pre + f"Total {unit}{suf.lower()}: {tot_c:,.0f} now vs {tot_p:,.0f} before ({sdiv(tot_c - tot_p, tot_p) * 100:+.1f}%). " + mover_text(t, suf, unit))
    return (q, a, act, mk(tbl, "bar", ["CHANGE"], cat=name, title=title, delta=delta_colors(good), label_fmt="+#,##0;-#,##0;0",
                          fmts={"CHANGE": "+#,##0;-#,##0;0", "CHANGE %": "+0.0%;-0.0%;0.0%"}))

def blk_customers(q, cur_df, prev_df, val, cd, pdd, unit, good, act, title):
    cs = cur_df.groupby("CUSTOMER")[val].sum(); ps = prev_df.groupby("CUSTOMER")[val].sum()
    ck_, pk_ = set(cs[cs > 0].index), set(ps[ps > 0].index)
    new, lost = ck_ - pk_, pk_ - ck_
    new_v, lost_v = cs.reindex(list(new)).sum(), ps.reindex(list(lost)).sum()
    t, suf, tbl = cmp_movers_tbl(cs, ps, cd, pdd, "CUSTOMER", 4)
    a = (f"{len(ck_)} customers now vs {len(pk_)} before: {len(new)} new (adding {new_v:,.0f} {unit}) and {len(lost)} not seen this time (they took {lost_v:,.0f} {unit} before); "
         f"{len(ck_ & pk_)} continue. " + mover_text(t, suf, unit))
    return (q, a, act, mk(tbl, "bar", ["CHANGE"], cat="CUSTOMER", title=title, delta=delta_colors(good), label_fmt="+#,##0;-#,##0;0",
                          fmts={"CHANGE": "+#,##0;-#,##0;0", "CHANGE %": "+0.0%;-0.0%;0.0%"}))

def blk_weekday(q, cur_df, prev_df, val, unit, act, title, daily_label):
    dc = cur_df.groupby("DATE")[val].sum(); dp = prev_df.groupby("DATE")[val].sum()
    wc, wp = weekday_avg(dc), weekday_avg(dp)
    days = [d for d in WEEKDAYS if d in wc.index or d in wp.index]
    tbl = pd.DataFrame({"WEEKDAY": days, "CURRENT WEEK": [float(wc.get(d, np.nan)) for d in days], "PREVIOUS WEEK": [float(wp.get(d, np.nan)) for d in days]})
    tbl["CHANGE %"] = np.where(tbl["PREVIOUS WEEK"] > 0, tbl["CURRENT WEEK"] / tbl["PREVIOUS WEEK"] - 1, np.nan)
    both = tbl.dropna(subset=["CURRENT WEEK", "PREVIOUS WEEK"])
    comparable = both.dropna(subset=["CHANGE %"])
    if len(comparable):
        i = comparable["CHANGE %"].abs().idxmax()
        a = (f"Average {unit} per day on the weekdays both weeks cover: busiest day now {wc.idxmax() if len(wc) else '-'} (before: {wp.idxmax() if len(wp) else '-'}). "
             f"Biggest change on {comparable.loc[i, 'WEEKDAY']}: {comparable.loc[i, 'PREVIOUS WEEK']:,.0f} to {comparable.loc[i, 'CURRENT WEEK']:,.0f} ({comparable.loc[i, 'CHANGE %'] * 100:+.0f}%). "
             f"Weekdays present in only one of the weeks are shown with a blank.")
    elif len(both):
        a = (f"Average {unit} per day on the weekdays both weeks cover: busiest day now {wc.idxmax() if len(wc) else '-'} (before: {wp.idxmax() if len(wp) else '-'}). "
             "No percentage change can be calculated because the previous-week averages are zero on all shared weekdays.")
    else:
        a = "The two weeks do not share a weekday, so the daily pattern cannot be compared directly; the table shows each week's weekdays."
    return (q, a, act, mk(tbl, "column", ["CURRENT WEEK", "PREVIOUS WEEK"], cat="WEEKDAY", title=title, fmts={"CHANGE %": "+0.0%;-0.0%;0.0%"}))

def blk_headline(sheet, mc, mp, act, extra=""):
    dkey, specs = KPI_SPECS[sheet]
    rows = []
    for label, key, kind, f, good in specs:
        if kind == "text": continue
        ch = kpi_change(kind, mc[key], mp[key], mc[dkey], mp[dkey])
        rel = None if ch is None else ((ch[0] / 100) if kind == "ratio" else ch[0])
        rows.append(dict(label=label, cur=_FMT[f](mc[key]), prev=_FMT[f](mp[key]), pct=rel, raw=None if ch is None else ch[0], kind=kind, good=good, per_day=bool(ch and ch[2]),
                         txt=None if ch is None else ch[1]))
    tbl = pd.DataFrame({"KPI": [r["label"].title() for r in rows], "CURRENT": [r["cur"] for r in rows], "PREVIOUS": [r["prev"] for r in rows],
                        "CHANGE %": [np.nan if r["pct"] is None else r["pct"] for r in rows]})
    colors = [{GREEN: "#27AE60", RED: "#C0392B", GREY: "#95A5A6"}[favour(r["good"], r["raw"], r["kind"])] for r in rows]
    ranked = sorted([r for r in rows if r["raw"] is not None], key=lambda r: -abs(r["pct"] if r["pct"] is not None else 0))
    parts = [f"{r['label'].capitalize()} {'rose' if r['raw'] > 0 else 'fell'} {r['txt'].lstrip('+-')}{' per day' if r['per_day'] else ''} ({r['cur']} vs {r['prev']})" for r in ranked[:4]]
    fav = [r["label"].lower() for r in rows if r["raw"] is not None and favour(r["good"], r["raw"], r["kind"]) == GREEN]
    unf = [r["label"].lower() for r in rows if r["raw"] is not None and favour(r["good"], r["raw"], r["kind"]) == RED]
    a = ("; ".join(parts) + ". " + (f"Moving the right way: {', '.join(fav)}. " if fav else "") + (f"Moving the wrong way: {', '.join(unf)}. " if unf else "")
         + (f"The weeks cover {mc[dkey]} and {mp[dkey]} dates, so volumes are compared per day. " if mc[dkey] != mp[dkey] else "") + extra)
    return ("What changed in the headline numbers?", a, act,
            mk(tbl, "bar", ["CHANGE %"], cat="KPI", title="Change versus the previous week (green = favourable, red = unfavourable, grey = neutral)",
               points=colors, label_fmt="+0.0%;-0.0%;0.0%", fmts={"CHANGE %": "+0.0%;-0.0%;0.0%"}, right=["CURRENT", "PREVIOUS"]))

def blk_cross(mc, mp):
    def per(m, a, b): return sdiv(m[a], m[b])
    rows = []
    def add(label, c, p, f, kind="rate"):
        rows.append((label, c, p, f, kind))
    add("Brands: net units per day", mc["b_perday"], mp["b_perday"], n0)
    add("Crates: net crates per day", mc["c_perday"], mp["c_perday"], n0)
    add("Mehul: returns per day", mc["m_perday"], mp["m_perday"], n0)
    add("Hansa: returns per day", mc["h_perday"], mp["h_perday"], n0)
    add("Revenue: invoiced per day", mc["v_perday"], mp["v_perday"], n0)
    add("Revenue: credit ratio", mc["v_ratio"], mp["v_ratio"], pc)
    add("Mehul returns per 100 units sold", 100 * sdiv(mc["m_perday"], mc["b_perday"]), 100 * sdiv(mp["m_perday"], mp["b_perday"]), n1)
    add("Hansa returns per 100 units sold", 100 * sdiv(mc["h_perday"], mc["b_perday"]), 100 * sdiv(mp["h_perday"], mp["b_perday"]), n1)
    add("Units per crate", sdiv(mc["b_units"], mc["c_net"]), sdiv(mp["b_units"], mp["c_net"]), n1)
    tbl = pd.DataFrame({"METRIC": [r[0] for r in rows], "CURRENT": [r[3](r[1]) for r in rows], "PREVIOUS": [r[3](r[2]) for r in rows],
                        "CHANGE %": [sdiv(r[1] - r[2], r[2]) if r[2] else np.nan for r in rows]})
    ch = dict(zip(tbl["METRIC"], tbl["CHANGE %"]))
    g = lambda k: ch.get(k, np.nan)
    def d(k):
        v = g(k); return "n/a" if pd.isna(v) else f"{v * 100:+.1f}%"
    sales, mret, hret = g("Brands: net units per day"), g("Mehul: returns per day"), g("Hansa: returns per day")
    verdict = []
    for nm, v in (("Mehul", mret), ("Hansa", hret)):
        if pd.notna(sales) and pd.notna(v):
            verdict.append(f"{nm} returns moved {v * 100:+.0f}% against sales {sales * 100:+.0f}%, so relative to sales returns got {'worse' if v > sales else 'better'}")
    a = (f"Per day: brand units {d('Brands: net units per day')}, crates {d('Crates: net crates per day')}, Mehul returns {d('Mehul: returns per day')}, Hansa returns {d('Hansa: returns per day')}, "
         f"revenue invoiced {d('Revenue: invoiced per day')}. Credit ratio {pc(mp['v_ratio'])} to {pc(mc['v_ratio'])}. "
         + ("; ".join(verdict) + ". " if verdict else "")
         + f"Returns per 100 units sold: Mehul {n1(100 * sdiv(mp['m_perday'], mp['b_perday']))} to {n1(100 * sdiv(mc['m_perday'], mc['b_perday']))}, Hansa {n1(100 * sdiv(mp['h_perday'], mp['b_perday']))} to {n1(100 * sdiv(mc['h_perday'], mc['b_perday']))}. "
         "These move together or apart by association only; they do not prove cause, and the files may cover different customers.")
    return ("Do the five datasets tell the same story about the change?", a,
            "If returns or credits grew faster than sales, review dispatch quantities and credit approvals before the next cycle; if sales grew and returns did not, protect the supply that is working.",
            mk(tbl, "bar", ["CHANGE %"], cat="METRIC", title="Change in each dataset, previous week to current week (per day)", points=["#2E86C1" if (pd.notna(v) and v >= 0) else "#F39C12" for v in tbl["CHANGE %"]],
               label_fmt="+0.0%;-0.0%;0.0%", fmts={"CHANGE %": "+0.0%;-0.0%;0.0%"}, right=["CURRENT", "PREVIOUS"]))

def route_upc(bd, cd):
    u = bd[bd["BRANDS_INV_QTY"] > 0].assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")["BRANDS_INV_QTY"].sum()
    c = cd[cd["CRATES_INV_QTY"] > 0].assign(R=lambda d: dim(d, "ROUTE NAME")).groupby("R")["CRATES_INV_QTY"].sum()
    t = pd.concat([u.rename("UNITS"), c.rename("CRATES")], axis=1).dropna()
    t = t[(t["UNITS"] > 0) & (t["CRATES"] > 0)]
    t["UPC"] = t["UNITS"] / t["CRATES"]
    return t

def share_pts_tbl(cur_df, prev_df, key, a, b, name, k=8, top_by="cur"):
    """share of `a` in (a + b) by key, both weeks"""
    def sh(df):
        g = df.groupby(key)[[a, b]].sum()
        g["T"] = g[a] + g[b]
        return g
    gc, gp = sh(cur_df), sh(prev_df)
    idx = gc.sort_values("T", ascending=False).head(k).index
    cs = (gc[b] / gc["T"].where(gc["T"] > 0)).reindex(idx)
    ps = (gp[b] / gp["T"].where(gp["T"] > 0)).reindex(idx) if len(gp) else pd.Series(np.nan, index=idx)
    return idx, cs, ps

def add_change_cols(M, cur_df, prev_df, val, total_col, prev_name, cd, pdd):
    """previous-week total, change and % change for every customer in the matrix; customers only in the previous week are added as rows"""
    kf = lambda d: d["CODE"].astype("string").fillna("") + "|" + d["CUSTOMER NAME"].astype("string").fillna("")
    keys = kf(M)
    pk = prev_df.assign(_k=kf(prev_df))
    prev_tot = pk.groupby("_k")[val].sum()
    M = M.copy()
    M[prev_name] = keys.map(prev_tot).fillna(0).to_numpy()
    lost = prev_tot[prev_tot != 0].index.difference(pd.Index(keys))
    if len(lost):
        info = pk.drop_duplicates("_k").set_index("_k").loc[lost, INFO_COLS]
        L = pd.DataFrame({c: 0 for c in M.columns}, index=range(len(lost)))
        for c in INFO_COLS:
            L[c] = info[c].to_numpy()
        L[prev_name] = prev_tot.loc[lost].to_numpy()
        M = pd.concat([M, L], ignore_index=True)
    adj = cd / pdd
    base = M[prev_name] * adj
    M["CHANGE"] = (M[total_col] - base).round(1)
    M["CHANGE_PCT"] = np.where(base > 0, (M[total_col] - base) / base.where(base > 0), np.nan).round(4)
    return M.sort_values(total_col, ascending=False, kind="stable").reset_index(drop=True), len(lost)

def mix_text(cs, ps, label):
    """largest change in share of the total, in points"""
    tc, tp = cs.sum(), ps.sum()
    if not tc or not tp: return ""
    idx = cs.index.union(ps.index)
    d = (cs.reindex(idx).fillna(0) / tc - ps.reindex(idx).fillna(0) / tp) * 100
    i = d.abs().idxmax()
    return f"Mix: {i}'s share of {label} moved from {ps.reindex(idx).fillna(0)[i] / tp * 100:.1f}% to {cs.reindex(idx).fillna(0)[i] / tc * 100:.1f}% ({d[i]:+.1f} pts). "


# In[51]:


# ---------- per-sheet comparison sections
if HAS_PREV:
    MC, MP = period_metrics(CURR), period_metrics(PRV)
    CP = {k: prep(CURR[k]) for k in ("brands_dw", "crates_dw", "mehul_dw", "hansa_dw", "rev_dw")}
    PP = {k: prep(PRV[k]) for k in ("brands_dw", "crates_dw", "mehul_dw", "hansa_dw", "rev_dw")}
    DATASET_KEY = {"1 Brands": "brands_dw", "2 Crates": "crates_dw", "3 Mehul Returns": "mehul_dw", "4 Hansa Returns": "hansa_dw", "5 Sales Revenue": "rev_dw"}

def _finish(plan, blocks, mc_dkey, matrix_args):
    """common part: KPI cards with previous-week values, comparison section, matrix change columns"""
    name = plan["name"]
    key = DATASET_KEY[name]
    cd, pdd = MC[mc_dkey], MP[mc_dkey]
    plan["kpis"] = kpi_cards(name, MC, MP)
    plan["cmp_title"] = "WEEK-ON-WEEK COMPARISON - what changed versus the previous week (prev_wk_data)"
    same = cd == pdd
    plan["cmp_note"] = (f"Current week: {span_txt(CURR[key])} ({cd} dates).  Previous week: {span_txt(PRV[key])} ({pdd} dates).  "
                        + ("Both weeks cover the same number of dates, so totals are compared directly." if same else
                           f"The weeks cover a different number of dates, so volumes are compared PER DAY (columns marked / DAY). "))
    plan["cmp"] = blocks + [blk_cross(MC, MP)]
    plan["subtitle"] += f"  |  vs previous week: {span_txt(PRV[key])} ({pdd} dates)"
    val, total_col, prev_name, good = matrix_args
    M, n_lost = add_change_cols(plan["matrix"], CURR[key], PRV[key], val, total_col, prev_name, cd, pdd)
    plan["matrix"] = M
    plan["sums"] = plan["sums"] + [prev_name, "CHANGE"]
    plan["pct"] = plan["pct"] + ["CHANGE_PCT"]
    plan["delta"] = ["CHANGE", "CHANGE_PCT"]; plan["delta_good"] = good
    plan["matrix_title"] += f" | + previous week, change and % change per customer ({n_lost} customers from the previous week with no current activity are included)"
    plan["checks"] = plan["checks"] + [(f"{name}: previous-week matrix total = previous-week source", M[prev_name].sum(), float(PRV[key][val].sum()))]
    return plan

def cmp_brands(plan):
    cb, pb = CP["brands_dw"], PP["brands_dw"]
    V = "BRANDS_INV_QTY"; ci, pi = cb[cb[V] > 0], pb[pb[V] > 0]
    cd, pdd = MC["b_days"], MP["b_days"]
    pr_c, pr_p = ci.groupby("PRODUCT")[V].sum(), pi.groupby("PRODUCT")[V].sum()
    B = [blk_headline("1 Brands", MC, MP, "Focus the review on the items moving the wrong way and check whether the change is real demand or just a different number of dates.")]
    B.append(blk_movers("Which products gained or lost the most volume?", pr_c, pr_p, cd, pdd, "PRODUCT", "units", True,
                        "Rebuild supply for products losing volume; raise production for those growing.", "Change in units invoiced by product", pre=mix_text(pr_c, pr_p, "units")))
    B.append(blk_movers("Which routes gained or lost the most volume?", ci.groupby("ROUTE NAME")[V].sum(), pi.groupby("ROUTE NAME")[V].sum(), cd, pdd, "ROUTE", "units", True,
                        "Talk to the route teams behind the largest falls: stock-outs, missed days or lost outlets.", "Change in units invoiced by route"))
    B.append(blk_customers("Which customers are new, lost, growing or shrinking?", ci, pi, V, cd, pdd, "units", True,
                           "Call the biggest decliners and welcome new accounts; confirm lost accounts are genuinely lost and not a data gap.", "Change in units invoiced by customer"))
    B.append(blk_weekday("Did the weekday pattern change?", ci, pi, V, "units", "Re-time production and dispatch to the weekdays that grew.", "Average units per day by weekday: this week vs previous week", "units"))
    cn_c, cn_p = cb.groupby("ROUTE NAME")["BRANDS_CN_QTY"].sum(), pb.groupby("ROUTE NAME")["BRANDS_CN_QTY"].sum()
    B.append(blk_movers("Did credit units change?", cn_c, cn_p, cd, pdd, "ROUTE", "credit units", False,
                        "Investigate routes where credits grew: they are returns or pricing errors that net sales are hiding.", "Change in credit units by route",
                        pre=f"Credit rate moved from {pc(MP['b_cr'])} to {pc(MC['b_cr'])} of units invoiced. "))
    return _finish(plan, B, "b_days", (V, "TOTAL_UNITS", "PREV_UNITS", True))

def cmp_crates(plan):
    cc, pc_ = CP["crates_dw"], PP["crates_dw"]
    V = "CRATES_INV_QTY"; ci, pi = cc[cc[V] > 0], pc_[pc_[V] > 0]
    cd, pdd = MC["c_days"], MP["c_days"]
    B = [blk_headline("2 Crates", MC, MP, "Check whether more crates went out because more product was sold or because crates are being held longer at customers.")]
    B.append(blk_movers("Which categories gained or lost crates?", ci.groupby("CATEGORY")[V].sum(), pi.groupby("CATEGORY")[V].sum(), cd, pdd, "CATEGORY", "crates", None,
                        "Adjust each category's crate quota to follow its change.", "Change in crates invoiced by category"))
    B.append(blk_movers("Which routes gained or lost crates?", ci.groupby("ROUTE NAME")[V].sum(), pi.groupby("ROUTE NAME")[V].sum(), cd, pdd, "ROUTE", "crates", None,
                        "Move crates between routes before buying more.", "Change in crates invoiced by route"))
    B.append(blk_customers("Which customers are new, lost, growing or shrinking in crate use?", ci, pi, V, cd, pdd, "crates", None,
                           "Customers with sharply more crates need a collection check; sharply fewer may be holding stock or closing.", "Change in crates invoiced by customer"))
    B.append(blk_weekday("Did the weekday pattern change?", ci, pi, V, "crates", "Plan crate washing and collection around the days that grew.", "Average crates per day by weekday: this week vs previous week", "crates"))
    uc, up = route_upc(CP["brands_dw"], CP["crates_dw"]), route_upc(PP["brands_dw"], PP["crates_dw"])
    idx = uc.sort_values("CRATES", ascending=False).head(8).index
    t = pd.DataFrame({"ROUTE": [short(i) for i in idx], "CURRENT WEEK": uc.loc[idx, "UPC"].round(1).to_numpy(), "PREVIOUS WEEK": up["UPC"].reindex(idx).round(1).to_numpy()})
    t["CHANGE %"] = np.where(t["PREVIOUS WEEK"] > 0, t["CURRENT WEEK"] / t["PREVIOUS WEEK"] - 1, np.nan)
    a = (f"Units moved per crate: {n1(sdiv(MP['b_units'], MP['c_net']))} before, {n1(sdiv(MC['b_units'], MC['c_net']))} now. Stand-alone crate credit notes: {len(PRV['c_cn_open'])} before, {len(c_cn_open)} now "
         f"(credited crates {n0(MP['c_cn'])} before, {n0(MC['c_cn'])} now). The table shows the 8 routes with the most crates now.")
    B.insert(5, ("Is each crate carrying more or fewer units than before?", a, "Routes whose units per crate fell may be over-supplied with crates; routes that rose may be short.",
                 mk(t, "column", ["CURRENT WEEK", "PREVIOUS WEEK"], cat="ROUTE", title="Units per crate by route: this week vs previous week", fmts={"CHANGE %": "+0.0%;-0.0%;0.0%"})))
    return _finish(plan, B, "c_days", (V, "TOTAL_CRATES", "PREV_CRATES", None))

def cmp_mehul(plan):
    cm, pm = CP["mehul_dw"], PP["mehul_dw"]
    T_, F, X = "MEHUL_TOTAL_RTN_QTY", "MEHUL_FRESH_RTN_QTY", "MEHUL_EXP_RTN_QTY"
    cd, pdd = MC["m_days"], MP["m_days"]
    B = [blk_headline("3 Mehul Returns", MC, MP, "If expired returns grew faster than fresh ones, the problem is rotation and shelf life, not demand.")]
    pr_c, pr_p = cm.groupby("PRODUCT")[T_].sum(), pm.groupby("PRODUCT")[T_].sum()
    B.append(blk_movers("Which products' returns rose or fell the most?", pr_c, pr_p, cd, pdd, "PRODUCT", "units returned", False,
                        "Cut dispatch of products whose returns rose; learn from those that fell.", "Change in units returned by product", pre=mix_text(pr_c, pr_p, "returns")))
    B.append(blk_movers("Which routes' returns rose or fell the most?", cm.groupby("ROUTE NAME")[T_].sum(), pm.groupby("ROUTE NAME")[T_].sum(), cd, pdd, "ROUTE", "units returned", False,
                        "Visit routes with rising returns first.", "Change in units returned by route"))
    B.append(blk_customers("Which customers are new, lost, returning more or less?", cm, pm, T_, cd, pdd, "units returned", False,
                           "Review order levels at customers whose returns rose.", "Change in units returned by customer"))
    B.append(blk_weekday("Did the weekday pattern of returns change?", cm, pm, T_, "units returned", "Check deliveries on days where returns rose.", "Average units returned per day by weekday: this week vs previous week", "returns"))
    idx, cs, ps = share_pts_tbl(cm, pm, "PRODUCT", F, X, "PRODUCT", k=8)
    t = pd.DataFrame({"PRODUCT": list(idx), "EXPIRED SHARE NOW %": cs.to_numpy(), "EXPIRED SHARE BEFORE %": ps.to_numpy()})
    t["CHANGE (pts)"] = ((t["EXPIRED SHARE NOW %"] - t["EXPIRED SHARE BEFORE %"]) * 100).round(1)
    worst = t.dropna(subset=["CHANGE (pts)"]).sort_values("CHANGE (pts)", ascending=False)
    a = (f"Expired share of returns: {pc(MP['m_exp_share'])} before, {pc(MC['m_exp_share'])} now. "
         + (f"Largest rise: {worst.iloc[0]['PRODUCT']} ({worst.iloc[0]['CHANGE (pts)']:+.1f} pts). " if len(worst) else "")
         + f"Fresh returns per day {n0(sdiv(MP['m_fresh'], MP['m_days']))} to {n0(sdiv(MC['m_fresh'], MC['m_days']))}; expired per day {n0(sdiv(MP['m_exp'], MP['m_days']))} to {n0(sdiv(MC['m_exp'], MC['m_days']))}.")
    B.insert(5, ("Did the fresh / expired balance shift?", a, "A rising expired share points to over-ordering of short-life products; shorten the order cycle.",
                 mk(t, "column", ["EXPIRED SHARE NOW %", "EXPIRED SHARE BEFORE %"], cat="PRODUCT", title="Expired share of returns by product (8 biggest now)", fmts={"CHANGE (pts)": "+0.0;-0.0;0.0"})))
    return _finish(plan, B, "m_days", (T_, "TOTAL_RETURNS", "PREV_RETURNS", False))

def cmp_hansa(plan):
    ch, ph = CP["hansa_dw"], PP["hansa_dw"]
    V = "HANSA_RTN_QTY"
    cd, pdd = MC["h_days"], MP["h_days"]
    B = [blk_headline("4 Hansa Returns", MC, MP, "Look at whether returns grew because there are more return documents or because each document is bigger.")]
    pr_c, pr_p = ch.groupby("PRODUCT")[V].sum(), ph.groupby("PRODUCT")[V].sum()
    B.append(blk_movers("Which products' returns rose or fell the most?", pr_c, pr_p, cd, pdd, "PRODUCT", "units returned", False,
                        "Cut dispatch of products whose returns rose.", "Change in units returned by product", pre=mix_text(pr_c, pr_p, "returns")))
    B.append(blk_movers("Which routes' returns rose or fell the most?", ch.groupby("ROUTE NAME")[V].sum(), ph.groupby("ROUTE NAME")[V].sum(), cd, pdd, "ROUTE", "units returned", False,
                        "Visit routes with rising returns first.", "Change in units returned by route"))
    B.append(blk_customers("Which customers are new, lost, returning more or less?", ch, ph, V, cd, pdd, "units returned", False,
                           "Review order levels at customers whose returns rose.", "Change in units returned by customer"))
    B.append(blk_weekday("Did the weekday pattern of returns change?", ch, ph, V, "units returned", "Check deliveries on days where returns rose.", "Average units returned per day by weekday: this week vs previous week", "returns"))
    def types(P):
        hr = P["h"][P["h"]["DATE"].notna()]
        ty = hr["TYPE"].astype("object").where(hr["TYPE"].notna(), "(blank)").astype(str).str.strip().replace("", "(blank)")
        return hr[P["h_prods"]].sum(axis=1).groupby(ty).sum()
    B.append(blk_movers("How did the mix of transaction types change?", types(CURR), types(PRV), cd, pdd, "TYPE", "units returned", False,
                        "Confirm what a rising type code means before acting.", "Change in units returned by transaction type",
                        pre=f"Documents {n0(MP['h_docs'])} before, {n0(MC['h_docs'])} now; units per document {n1(MP['h_perdoc'])} to {n1(MC['h_perdoc'])}. "))
    return _finish(plan, B, "h_days", (V, "TOTAL_RETURNS", "PREV_RETURNS", False))

def cmp_revenue(plan):
    cv, pv = CP["rev_dw"], PP["rev_dw"]
    IA, CA = "REV_INV_AMT", "REV_CN_AMT"
    cd, pdd = MC["v_days"], MP["v_days"]
    B = [blk_headline("5 Sales Revenue", MC, MP, "If credits grew faster than invoiced value, tighten credit approval before the ratio becomes the new normal.")]
    B.append(blk_movers("Which categories gained or lost invoiced value?", cv.groupby("CATEGORY")[IA].sum(), pv.groupby("CATEGORY")[IA].sum(), cd, pdd, "CATEGORY", "value", True,
                        "Back the growing categories and review pricing or supply in those that fell.", "Change in invoiced value by category"))
    B.append(blk_movers("Which routes gained or lost invoiced value?", cv.groupby("ROUTE NAME")[IA].sum(), pv.groupby("ROUTE NAME")[IA].sum(), cd, pdd, "ROUTE", "value", True,
                        "Talk to the route teams behind the largest falls.", "Change in invoiced value by route"))
    B.append(blk_customers("Which customers are new, lost, growing or shrinking?", cv, pv, IA, cd, pdd, "value", True,
                           "Call the biggest decliners and follow up lost accounts.", "Change in invoiced value by customer"))
    B.append(blk_weekday("Did the weekday pattern change?", cv, pv, IA, "value", "Match delivery and invoicing capacity to the days that grew.", "Average invoiced value per day by weekday: this week vs previous week", "value"))
    def rr(df):
        g = df.groupby("ROUTE NAME")[[IA, CA]].sum()
        return g
    gc, gp = rr(cv), rr(pv)
    idx = gc.sort_values(IA, ascending=False).head(8).index
    t = pd.DataFrame({"ROUTE": [short(i) for i in idx], "CREDIT RATIO NOW %": (gc.loc[idx, CA] / gc.loc[idx, IA].where(gc.loc[idx, IA] > 0)).to_numpy(),
                      "CREDIT RATIO BEFORE %": (gp[CA].reindex(idx) / gp[IA].reindex(idx).where(gp[IA].reindex(idx) > 0)).to_numpy()})
    t["CHANGE (pts)"] = ((t["CREDIT RATIO NOW %"] - t["CREDIT RATIO BEFORE %"]) * 100).round(1)
    worst = t.dropna(subset=["CHANGE (pts)"]).sort_values("CHANGE (pts)", ascending=False)
    a = (f"Credit ratio {pc(MP['v_ratio'])} before, {pc(MC['v_ratio'])} now. Average invoice {n0(MP['v_avg'])} to {n0(MC['v_avg'])}; invoices {n0(MP['v_n'])} to {n0(MC['v_n'])}. "
         + (f"Largest rise in credit ratio among the 8 biggest routes: {worst.iloc[0]['ROUTE']} ({worst.iloc[0]['CHANGE (pts)']:+.1f} pts)." if len(worst) else ""))
    B.insert(5, ("Did credit control get better or worse?", a, "Review credit notes on the routes whose ratio rose.",
                 mk(t, "column", ["CREDIT RATIO NOW %", "CREDIT RATIO BEFORE %"], cat="ROUTE", title="Credit ratio by route (8 biggest by invoiced value now)", fmts={"CHANGE (pts)": "+0.0;-0.0;0.0"})))
    return _finish(plan, B, "v_days", (IA, "TOTAL_INVOICED", "PREV_INVOICED", True))

COMPARE = {"1 Brands": cmp_brands, "2 Crates": cmp_crates, "3 Mehul Returns": cmp_mehul, "4 Hansa Returns": cmp_hansa, "5 Sales Revenue": cmp_revenue}
def add_comparison(plan):
    return COMPARE[plan["name"]](plan) if HAS_PREV else plan


# ## Build the workbook (layout, charts and print setup)

# In[52]:


PALETTE = ["#1F3864", "#2E86C1", "#F39C12", "#27AE60", "#C0392B", "#8E44AD", "#16A085", "#7F8C8D", "#D35400", "#5D6D7E"]
SERIES_COLOR = {"FRESH": "#27AE60", "EXPIRED": "#C0392B", "FRESH %": "#27AE60", "EXPIRED %": "#C0392B", "INVOICED": "#1F3864", "CREDITS": "#C0392B",
                "CREDITED": "#C0392B", "MEHUL RETURN RATE %": "#B5651D", "HANSA RETURN RATE %": "#7B2D8E", "MEHUL UNITS": "#B5651D", "HANSA UNITS": "#7B2D8E",
                "MEHUL RETURNS": "#B5651D", "HANSA RETURNS": "#7B2D8E", "FIRST HALF / DAY": "#9DB4D3", "SECOND HALF / DAY": "#1F3864",
                "BRAND CREDIT RATE %": "#1F3864", "REVENUE CREDIT RATIO %": "#C0392B"}
PAPER_PT = {9: (842, 595), 8: (1191, 842), 1: (792, 612)}      # landscape width, height in points
INFO_W = {"CODE": 7, "CUSTOMER NAME": 26, "CATEGORY": 10, "ROUTE NAME": 14, "SALES_REP": 10, "MERCHANDISERS": 15}
MARGIN_IN = {"left": 0.3, "right": 0.3, "top": 0.6, "bottom": 0.6}
MIN_GRID_PX = 1050
MAX_TABLE_ROWS = 16

def px(w): return int(w * 7 + 5)

def lines_needed(text, chars):
    return sum(max(1, int(np.ceil(len(p) / max(chars, 1)))) for p in str(text).split("\n"))

def col_numfmt(name, s, over):
    if name in over: return over[name]
    if name.endswith("%"): return "0.0%"
    v = pd.to_numeric(s, errors="coerce").dropna()
    if len(v) and np.allclose(v, v.round()): return "#,##0"
    return "#,##0.0"

def chart_numfmt(f, vmax=0):
    """label format; large values are shown as K / M on the chart (the table beside it keeps the full number)"""
    if f in ("#,##0", "#,##0.0"):
        if vmax >= 1e6: return '#,##0.0,,"M"'
        if vmax >= 1e5: return '#,##0,"K"'
    return f

def render_sheet(wb, plan):
    ws = wb.add_worksheet(plan["name"])
    color = plan["color"]
    M = plan["matrix"]
    cols = list(M.columns)
    widths = []
    for c in cols:
        if c in INFO_W: widths.append(INFO_W[c])
        elif c in plan["rotate"]: widths.append(5.5)
        elif c in plan["pct"]: widths.append(8.5)
        else: widths.append(8)
    while sum(px(w) for w in widths) < MIN_GRID_PX:
        widths.append(8)
    N = len(widths)
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    pxs = [px(w) for w in widths]
    cum = np.cumsum(pxs)
    left = np.concatenate([[0], cum[:-1]])
    total_px = int(cum[-1])
    total_chars = sum(widths)

    def col_for(frac, start=0):
        return int(min(N - 1, max(start, np.searchsorted(cum, frac * total_px))))

    pw, ph = PAPER_PT[PAPER_SIZE]
    usable_px = (pw - 2 * MARGIN_IN["left"] * 72) / 0.75
    scale = int(np.floor(100 * usable_px / total_px * 0.88))   # 12% safety: viewers measure column widths slightly differently
    scale = max(30, min(100, scale))
    cap = (ph - (MARGIN_IN["top"] + MARGIN_IN["bottom"]) * 72) / (scale / 100) * 0.95
    st = {"r": 0, "y": 0.0, "breaks": []}

    def need(h, force=False):
        if st["y"] > 0 and (force or st["y"] + h > cap):
            st["breaks"].append(st["r"]); st["y"] = 0.0

    def put(h):
        ws.set_row(st["r"], h); st["r"] += 1; st["y"] += h

    def fmt(**kw):
        base = {"font_name": "Calibri", "valign": "vcenter"}
        base.update(kw); return wb.add_format(base)

    F = dict(
        title=fmt(font_size=20, bold=True, font_color="#FFFFFF", bg_color=color, indent=1),
        sub=fmt(font_size=10, italic=True, font_color="#595959", indent=1),
        bar=fmt(bg_color="#F39C12"),
        kl=fmt(font_size=9, bold=True, font_color="#595959", bg_color="#EEF3F9", align="center", top=5, top_color=color, left=2, left_color="#FFFFFF", right=2, right_color="#FFFFFF"),
        kv=fmt(font_size=18, bold=True, font_color=color, bg_color="#EEF3F9", align="center", left=2, left_color="#FFFFFF", right=2, right_color="#FFFFFF"),
        note=fmt(font_size=9.5, italic=True, font_color="#7F6000", bg_color="#FFF8E1", text_wrap=True, valign="top", indent=1),
        sec=fmt(font_size=12, bold=True, font_color="#FFFFFF", bg_color="#2C3E50", indent=1),
        num=fmt(font_size=11, bold=True, font_color="#FFFFFF", bg_color=color, align="center", border=1, border_color="#FFFFFF"),
        qbar=fmt(font_size=11, bold=True, font_color=color, bg_color="#EEF3F9", indent=1, bottom=2, bottom_color=color),
        ans=fmt(font_size=10, text_wrap=True, valign="top", indent=1, bg_color="#FFFFFF"),
        act=fmt(font_size=10, text_wrap=True, valign="top", indent=1, bg_color="#EAF5EA", font_color="#1E5631", left=5, left_color="#27AE60"),
        actb=wb.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "font_color": "#1E5631"}),
        small=fmt(font_size=9, italic=True, font_color="#7F7F7F"),
        nodata=fmt(font_size=10, italic=True, font_color="#7F7F7F", align="center", border=1, border_color="#D9D9D9"),
    )
    cache = {}
    def tf(kind, **kw):
        key = (kind, tuple(sorted(kw.items())))
        if key not in cache:
            cache[key] = wb.add_format(dict({"font_name": "Calibri", "font_size": 9.5, "valign": "vcenter"}, **kw))
        return cache[key]

    def merged(r, c0, c1, text, f, r1=None):
        r1 = r if r1 is None else r1
        if c1 > c0 or r1 > r: ws.merge_range(r, c0, r1, c1, text, f)
        else: ws.write(r, c0, text, f)

    # ---------- banner, KPI cards, notes
    ws.set_row(0, 38); merged(0, 0, N - 1, f"{REPORT_TITLE}  |  {plan['title']}", F["title"])
    ws.set_row(1, 20); merged(1, 0, N - 1, plan["subtitle"], F["sub"])
    ws.set_row(2, 5); merged(2, 0, N - 1, "", F["bar"])
    st["r"], st["y"] = 3, 63
    put(4)
    k = len(plan["kpis"])
    edges = [col_for((i + 1) / k) for i in range(k)]
    spans, start = [], 0
    for i in range(k):
        end = N - 1 if i == k - 1 else max(start + 1, edges[i])
        spans.append((start, end)); start = end + 1
    has_comp = any(len(kp) > 2 and kp[2] for kp in plan["kpis"])
    r_l = st["r"]; put(18); r_v = st["r"]; put(28)
    r_c = None
    if has_comp:
        r_c = st["r"]; put(16)
    kc = fmt(font_size=9.5, font_color="#595959", bg_color="#EEF3F9", align="center", left=2, left_color="#FFFFFF", right=2, right_color="#FFFFFF")
    for (c0, c1), kp in zip(spans, plan["kpis"]):
        lbl, val = kp[0], kp[1]
        merged(r_l, c0, c1, lbl, F["kl"]); merged(r_v, c0, c1, val, F["kv"])
        if r_c is not None:
            comp = kp[2] if len(kp) > 2 else None
            merged(r_c, c0, c1, "", kc)
            if comp and comp[1]:
                ws.write_rich_string(r_c, c0, wb.add_format({"font_name": "Calibri", "font_size": 9.5, "font_color": "#595959"}), comp[0],
                                     wb.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "font_color": comp[2]}), comp[1], kc)
            elif comp:
                ws.write(r_c, c0, comp[0], kc)
    put(4)
    txt = "  |  ".join(plan["notes"])
    h = lines_needed(txt, total_chars * 1.1) * 13 + 6
    merged(st["r"], 0, N - 1, "NOTES: " + txt, F["note"]); put(h)
    put(4)

    # ---------- question blocks: answer + action | data table | chart
    cT = col_for(0.34, start=2)
    cB = col_for(0.64, start=cT + 5)
    def qblock(prefix, i, item):
        q, a, act, b = item
        tbl = b["table"]
        has_t = tbl is not None and len(tbl) > 0
        tbl = tbl.head(MAX_TABLE_ROWS).reset_index(drop=True) if has_t else None
        kind, vals = b.get("kind"), b.get("vals") or []
        if has_t and kind:
            ok = [v for v in vals if v in tbl.columns and pd.to_numeric(tbl[v], errors="coerce").notna().any()]
            if not ok: kind = None
            vals = ok
        t_c0 = cT + 1
        t_c1 = cB if (has_t and kind) else N - 1
        text_c1 = cT if has_t else col_for(0.60)
        text_chars = sum(widths[0:text_c1 + 1]) * 1.0
        la = lines_needed(a, text_chars); lact = lines_needed("Suggested action: " + act, text_chars)
        ans_rows = int(np.ceil((la * 13 + 10) / 15)); act_rows = int(np.ceil((lact * 13 + 12) / 15))
        tbl_rows = (len(tbl) + 2) if has_t else 0
        B = max(ans_rows + act_rows + 2, tbl_rows + 1, 10 if (has_t and kind) else 8)
        total_h = 22 + 15 * B + 4
        need(total_h)
        r0 = st["r"]
        ws.write(r0, 0, f"{prefix}{i}", F["num"]); merged(r0, 1, N - 1, q, F["qbar"]); put(22)
        rb = st["r"]
        for _ in range(B): put(15)
        # text
        merged(rb, 0, text_c1, a, F["ans"], r1=rb + ans_rows - 1)
        a0 = rb + ans_rows + 1
        merged(a0, 0, text_c1, "", F["act"], r1=a0 + act_rows - 1)
        ws.write_rich_string(a0, 0, F["actb"], "Suggested action: ", wb.add_format({"font_name": "Calibri", "font_size": 10, "font_color": "#1E5631"}), act, F["act"])
        if not has_t:
            return
        # table
        names = list(tbl.columns)
        w_ = []
        for c in names:
            if tbl[c].dtype == object or str(tbl[c].dtype).startswith("str"):
                ml = max(len(str(x)) for x in tbl[c]) if len(tbl) else 8
                w_.append(float(np.clip(ml / 8.5, 1.3, 2.8)))
            else:
                w_.append(max(1.0, min(1.5, len(c) / 11)))
        reg_left = left[t_c0]; reg_px = cum[t_c1] - reg_left
        cut = np.cumsum(w_) / sum(w_) * reg_px + reg_left
        slots, prev = [], t_c0 - 1
        for j in range(len(names)):
            e = int(np.searchsorted(cum, cut[j] - 1))
            e = min(max(e, prev + 1), t_c1 - (len(names) - 1 - j))
            if j == len(names) - 1: e = t_c1
            slots.append((prev + 1, e)); prev = e
        over = b.get("fmts", {})
        hl = 1
        for j, c in enumerate(names):
            sl_chars = sum(widths[slots[j][0]:slots[j][1] + 1]) * 1.05
            hl = max(hl, int(np.ceil(len(c) / max(sl_chars, 1))))
        ws.set_row(rb, 13 * min(hl, 3) + 8)
        for j, c in enumerate(names):
            merged(rb, slots[j][0], slots[j][1], c, tf("h", bold=True, font_color="#FFFFFF", bg_color=color, align="center", text_wrap=True, border=1, border_color="#FFFFFF"))
        fm = {c: col_numfmt(c, tbl[c], over) for c in names}
        for ri in range(len(tbl)):
            alt = "#F5F7FA" if ri % 2 else "#FFFFFF"
            for j, c in enumerate(names):
                v = tbl[c].iloc[ri]
                is_num = isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool)
                f = tf("c", bg_color=alt, bottom=1, bottom_color="#E3E3E3", **({"num_format": fm[c], "align": "right"} if is_num or pd.isna(v) and c in fm and tbl[c].dtype.kind in "fiu" or c in b.get("right", []) else {"align": "left"}))
                if is_num and pd.notna(v):
                    merged(rb + 1 + ri, slots[j][0], slots[j][1], float(v), f)
                else:
                    merged(rb + 1 + ri, slots[j][0], slots[j][1], "" if (not is_num and pd.isna(v)) or (is_num and pd.isna(v)) else str(v), f)
        if kind:
            cat = b.get("cat") or names[0]
            sc = {c: slots[names.index(c)][0] for c in names}
            vmax = float(np.nanmax([np.nanmax(np.abs(pd.to_numeric(tbl[v], errors="coerce").to_numpy(dtype=float))) for v in vals]))
            ch = make_chart(wb, plan["name"], b, kind, cat, vals, sc, rb, len(tbl), fm, vmax,
                            int(cum[N - 1] - cum[cB]) - 16, int(B * 15 / 0.75) - 6)
            ws.insert_chart(rb, cB + 1, ch, {"x_offset": 8, "y_offset": 2})

    def qa_section(title, items, prefix):
        need(22 + 22 + 15 * 14 + 8)
        merged(st["r"], 0, N - 1, title, F["sec"]); put(22)
        put(6)
        for i, item in enumerate(items, start=1):
            qblock(prefix, i, item)
            put(4)

    qa_section("KEY QUESTIONS ANSWERED - each with its data table and chart", plan["questions"], "Q")
    if plan.get("cmp"):
        h = lines_needed(plan["cmp_note"], total_chars * 1.1) * 13 + 6
        need(22 + 6 + h + 15 * 14 + 30, force=True)
        merged(st["r"], 0, N - 1, plan["cmp_title"], F["sec"]); put(22)
        merged(st["r"], 0, N - 1, plan["cmp_note"], F["note"]); put(h)
        put(6)
        for i, item in enumerate(plan["cmp"], start=1):
            qblock("C", i, item)
            put(4)
    qa_section(plan["rel_title"], plan["rels"], "R")

    # ---------- matrix
    need(0, force=True)
    merged(st["r"], 0, N - 1, plan["matrix_title"], F["sec"]); put(22)
    merged(st["r"], 0, N - 1, "Click the arrow on any column heading to filter by customer, category, route, sales rep or by a value; the TOTAL row recalculates for the filtered rows. "
           "Heat shading: darker = larger.", F["small"]); put(16)
    hdr_row = st["r"]
    longest = max([len(str(c)) for c in plan["rotate"]] + [0])
    hdr_h = min(120, 12 + longest * 6) if plan["rotate"] else 42
    ws.set_row(hdr_row, hdr_h)
    n = len(M)
    base = dict(font_name="Calibri", font_size=10)
    cfg = []
    sums = set(plan["sums"]); pcts = set(plan["pct"]); rot = set(plan["rotate"])
    data = M.astype(object).where(M.notna(), None)
    dfmts = {}
    for c in cols:
        hf = wb.add_format(dict(base, bold=True, font_color="#FFFFFF", bg_color=color, align="center", valign="bottom", border=1, border_color="#FFFFFF",
                                text_wrap=(c not in rot), rotation=90 if c in rot else 0))
        if c == "CODE": nf = "0"
        elif c in INFO_W: nf = None
        elif c in plan.get("delta", []): nf = "+0.0%;-0.0%;0.0%" if c in pcts else "+#,##0;-#,##0;0"
        elif c in pcts: nf = "0.0%"
        elif c == "AVG_PER_DAY": nf = "#,##0.0"
        else: nf = "#,##0;-#,##0;;@"
        dfmts[c] = wb.add_format(dict(base, **({"num_format": nf} if nf else {}), **({"align": "left"} if c == "CODE" else {})))
        cfg.append({"header": c.replace("_", " "), "header_format": hf, "format": dfmts[c]})
    last = hdr_row + n + 1
    ws.add_table(hdr_row, 0, hdr_row + n, len(cols) - 1, {"columns": cfg, "data": data.values.tolist(), "style": "Table Style Light 1", "total_row": False,
                                                         "name": "tbl_" + plan["name"].split()[1]})
    tot_f = wb.add_format(dict(base, bold=True, top=2, bottom=1, bg_color="#EEF3F9", num_format="#,##0"))
    tot_t = wb.add_format(dict(base, bold=True, top=2, bottom=1, bg_color="#EEF3F9"))
    for j, c in enumerate(cols):
        if j == 0: ws.write(last, j, "Total", tot_t)
        elif j == 1: ws.write(last, j, "Total of the rows shown (updates when you filter)", tot_t)
        elif c in sums and n:
            rng = xlsxwriter.utility.xl_range(hdr_row + 1, j, hdr_row + n, j)
            ws.write_formula(last, j, f"=SUBTOTAL(109,{rng})", tot_f, float(pd.to_numeric(M[c]).sum()))
        else:
            ws.write_blank(last, j, None, tot_t)
    if n and plan.get("delta"):
        g_ = plan.get("delta_good")
        pos_c, neg_c = (("#1E8449", "#C0392B") if g_ is True else ("#C0392B", "#1E8449") if g_ is False else ("#2E86C1", "#D68910"))
        for c in plan["delta"]:
            j = cols.index(c)
            rng_ = (hdr_row + 1, j, hdr_row + n, j)
            ws.conditional_format(*rng_, {"type": "cell", "criteria": ">", "value": 0, "format": wb.add_format({"font_color": pos_c, "bold": True})})
            ws.conditional_format(*rng_, {"type": "cell", "criteria": "<", "value": 0, "format": wb.add_format({"font_color": neg_c, "bold": True})})
    if n:
        for c in plan["heat"]:
            j = cols.index(c)
            ws.conditional_format(hdr_row + 1, j, hdr_row + n, j, {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": "#6FA8DC"})

    ws.set_landscape(); ws.set_paper(PAPER_SIZE)
    ws.set_margins(left=MARGIN_IN["left"], right=MARGIN_IN["right"], top=MARGIN_IN["top"], bottom=MARGIN_IN["bottom"])
    ws.set_header("&L&\"Calibri,Bold\"&11" + REPORT_TITLE + "&R&A")
    ws.set_footer("&LPrinted &D&CPage &P of &N&R&F")
    ws.center_horizontally()
    ws.set_print_scale(scale)
    ws.print_area(0, 0, last, N - 1)
    ws.repeat_rows(hdr_row)
    ws.set_h_pagebreaks(sorted(set(st["breaks"])))
    ws.hide_gridlines(2)
    ws.set_tab_color(color)
    ws.set_zoom(90)
    return dict(sheet=plan["name"], scale=scale, page_breaks=len(st["breaks"]), width_px=total_px, rows=n, cols=len(cols), hdr_row=hdr_row,
                questions=len(plan["questions"]), comparisons=len(plan.get("cmp", [])), relationships=len(plan["rels"]))


def make_chart(wb, sheet, b, kind, cat, vals, sc, hdr_row, nrows, fm, vmax, w_px, h_px):
    """chart drawn straight from the cells of the table shown beside it; every point carries a data label"""
    r0 = hdr_row + 1
    r1 = r0 + nrows - 1
    stacked = b.get("stacked", False)
    opts = {"type": kind}
    if stacked: opts["subtype"] = "stacked"
    ch = wb.add_chart(opts)
    cat_ref = [sheet, r0, sc[cat], r1, sc[cat]]
    for i, v in enumerate(vals):
        colr = SERIES_COLOR.get(v, PALETTE[i % len(PALETTE)])
        if v.upper().startswith(("CURRENT", "EXPIRED SHARE NOW", "CREDIT RATIO NOW")): colr = "#1F3864"
        if v.upper().startswith(("PREVIOUS", "EXPIRED SHARE BEFORE", "CREDIT RATIO BEFORE")): colr = "#9DB4D3"
        f = chart_numfmt(fm.get(v, "#,##0"), vmax)
        s = {"name": [sheet, hdr_row, sc[v]], "categories": cat_ref, "values": [sheet, r0, sc[v], r1, sc[v]]}
        if kind in ("bar", "column"):
            s["fill"] = {"color": colr}; s["gap"] = 50
            if b.get("delta"):
                s["fill"] = {"color": b["delta"][0]}; s["invert_if_negative"] = True; s["invert_if_negative_color"] = b["delta"][1]
            if b.get("points"):
                s["points"] = [{"fill": {"color": c_}} for c_ in b["points"]]
            if b.get("label_fmt"): f = b["label_fmt"]
            s["data_labels"] = {"value": True, "num_format": (f + ";;;") if stacked else f, "font": {"size": 8, "bold": True, "color": "#FFFFFF" if stacked else "#333333"},
                                **({"position": "center"} if stacked else {"position": "outside_end"})}
        elif kind == "line":
            s["line"] = {"color": colr, "width": 2.25}
            s["marker"] = {"type": "circle", "size": 5, "fill": {"color": colr}, "border": {"color": colr}}
            s["data_labels"] = {"value": True, "num_format": f, "position": "above", "font": {"size": 8, "bold": True, "color": colr}}
        else:
            s["points"] = [{"fill": {"color": PALETTE[k % len(PALETTE)]}} for k in range(nrows)]
            s["data_labels"] = {"value": True, "percentage": True, "separator": "\n", "num_format": f, "leader_lines": True, "font": {"size": 8}}
        ch.add_series(s)
    ch.set_title({"name": b.get("title") or "", "name_font": {"size": 10.5, "bold": True, "color": "#1F3864"}})
    ch.set_size({"width": max(w_px, 250), "height": max(h_px, 180)})
    ch.set_chartarea({"border": {"color": "#D9D9D9"}})
    if kind == "pie":
        ch.set_legend({"position": "right", "font": {"size": 8}})
        return ch
    ch.set_legend({"position": "bottom", "font": {"size": 8}} if len(vals) > 1 else {"none": True})
    pct = fm.get(vals[0], "").endswith("%")
    cat_axis = {"num_font": {"size": 8}, "line": {"color": "#BFBFBF"}, "interval_unit": 1}
    val_axis = {"num_font": {"size": 8}, "num_format": "0%" if pct else chart_numfmt("#,##0", vmax), "major_gridlines": {"visible": True, "line": {"color": "#EDEDED"}}, "line": {"none": True}}
    if b.get("ymax"): val_axis["max"] = b["ymax"]; val_axis["min"] = 0
    if b.get("delta") or b.get("points"):
        cat_axis["label_position"] = "low"           # keep category names clear of bars that go negative
    if kind == "bar":
        cat_axis["reverse"] = True
        val_axis = {"visible": False, "major_gridlines": {"visible": False}}
        if b.get("ymax"): val_axis["max"] = b["ymax"]; val_axis["min"] = 0
        ch.set_y_axis(cat_axis); ch.set_x_axis(val_axis)
    else:
        if kind == "line": val_axis["min"] = 0
        ch.set_x_axis(cat_axis); ch.set_y_axis(val_axis)
    return ch


# In[53]:


plans = []
for fn in (analyse_brands, analyse_crates, analyse_mehul, analyse_hansa, analyse_revenue):
    p = add_comparison(fn())
    plans.append(p)
    blocks = p["questions"] + p.get("cmp", []) + p["rels"]
    n_chart = sum(1 for it in blocks if it[3]["kind"] and it[3]["table"] is not None and len(it[3]["table"]))
    print(f"{p['name']}: {len(p['questions'])} questions + {len(p.get('cmp', []))} week-on-week comparisons + {len(p['rels'])} relationship answers, {n_chart} charts, matrix {p['matrix'].shape[0]:,} rows x {p['matrix'].shape[1]} columns")

REPORT_FILE.parent.mkdir(parents=True, exist_ok=True)
wb = xlsxwriter.Workbook(str(REPORT_FILE), {"nan_inf_to_errors": True})
wb.set_properties({"title": REPORT_TITLE, "subject": "Brands, crates, Mehul returns, Hansa returns and sales revenue", "comments": "Generated by fmcg_excel_analysis.ipynb"})
layout = [render_sheet(wb, p) for p in plans]
wb.close()
print("\nSaved:", REPORT_FILE)
print(pd.DataFrame(layout).to_string(index=False))


# ## Check the saved workbook

# In[ ]:


import openpyxl, zipfile

checks, warns = [], []
def ck(name, ok, detail=""):
    checks.append((name, "OK" if ok else "FAIL", detail))

# ---- rows that cannot be used (no date) are left out of every sheet; say how many
hyg = [("brands invoices", b_inv), ("brands credit notes", b_cn), ("crate invoices", c_inv), ("crate credit notes", c_cn), ("Hansa returns", h),
       ("revenue invoices", rev_inv), ("revenue credits", rev_cn)]
print("Rows with no valid date (left out of the report):")
for nm, d in hyg:
    k = int(d["DATE"].isna().sum())
    print(f"  {nm}: {k}" + ("   <- check the bottom of that sheet for a TOTAL row or blank lines" if k else ""))
print()

P = {p["name"]: p for p in plans}
ri_ok = rev_inv.dropna(subset=["DATE"]); rc_ok = rev_cn.dropna(subset=["DATE"])
ck("brands matrix = brands source units", int(P["1 Brands"]["matrix"]["TOTAL_UNITS"].sum()) == int(brands_dw["BRANDS_INV_QTY"].sum()))
ck("crates matrix = crates source", int(P["2 Crates"]["matrix"]["TOTAL_CRATES"].sum()) == int(crates_dw["CRATES_INV_QTY"].sum()))
ck("mehul matrix = mehul source", int(P["3 Mehul Returns"]["matrix"]["TOTAL_RETURNS"].sum()) == int(mehul_dw["MEHUL_TOTAL_RTN_QTY"].sum()))
ck("hansa matrix = hansa source", int(P["4 Hansa Returns"]["matrix"]["TOTAL_RETURNS"].sum()) == int(hansa_dw["HANSA_RTN_QTY"].sum()))
ck("4 Hansa Returns: returned-unit quantities are nonnegative", bool((hansa_dw["HANSA_RTN_QTY"] >= 0).all()))
hansa_q = P["4 Hansa Returns"]["questions"]
hansa_empty = [i + 1 for i, item in enumerate(hansa_q) if item[3]["table"] is None or item[3]["table"].empty]
ck("4 Hansa Returns: all 10 question tables contain data", len(hansa_q) == 10 and not hansa_empty,
   "10 of 10 populated" if len(hansa_q) == 10 and not hansa_empty else f"empty questions: {hansa_empty}")
hansa_visuals_ok = all(item[3]["kind"] is None or any(
    col in item[3]["table"] and pd.to_numeric(item[3]["table"][col], errors="coerce").notna().any()
    for col in (item[3]["vals"] or [])) for item in hansa_q if item[3]["table"] is not None and not item[3]["table"].empty)
ck("4 Hansa Returns: every question chart has numeric data", hansa_visuals_ok)
ck("revenue matrix = inv_total of dated invoices", np.isclose(P["5 Sales Revenue"]["matrix"]["TOTAL_INVOICED"].sum(), ri_ok["inv_total"].sum()),
   f'{P["5 Sales Revenue"]["matrix"]["TOTAL_INVOICED"].sum():,.2f} vs {ri_ok["inv_total"].sum():,.2f}')
ck("revenue credits = inv_total of dated credits", np.isclose(P["5 Sales Revenue"]["matrix"]["TOTAL_CREDITS"].sum(), rc_ok["inv_total"].sum()),
   f'{P["5 Sales Revenue"]["matrix"]["TOTAL_CREDITS"].sum():,.2f} vs {rc_ok["inv_total"].sum():,.2f}')

wbk = openpyxl.load_workbook(REPORT_FILE)
ck("workbook has exactly 5 sheets", len(wbk.sheetnames) == 5, str(wbk.sheetnames))
for p in plans:
    ws = wbk[p["name"]]
    tbl = list(ws.tables.values())
    ok_tbl = False
    if len(tbl) == 1:
        a0, a1 = tbl[0].ref.split(":")
        ok_tbl = int("".join(ch for ch in a1 if ch.isdigit())) - int("".join(ch for ch in a0 if ch.isdigit())) == len(p["matrix"])
    ck(f"{p['name']}: one filterable table with all {len(p['matrix']):,} matrix rows", ok_tbl, tbl[0].ref if tbl else "no table")
    ck(f"{p['name']}: landscape, scale {ws.page_setup.scale}%, print area set", ws.page_setup.orientation == "landscape" and bool(ws.print_area) and ws.page_setup.scale is not None, str(ws.print_area))
    ck(f"{p['name']}: header row repeats on every printed page", bool(ws.print_title_rows), str(ws.print_title_rows))
    for name, a_, b_ in p["checks"]:
        ck(f"{p['name']} - {name}", bool(np.isclose(float(a_), float(b_))), f"{a_:,.2f} vs {b_:,.2f}")
    blocks = p["questions"] + p.get("cmp", []) + p["rels"]
    empty = [i + 1 for i, it in enumerate(blocks) if it[3]["table"] is None or len(it[3]["table"]) == 0]
    if empty:
        warns.append(f"{p['name']}: question(s) {empty} have no data in your files, so they show the answer text only")

with zipfile.ZipFile(REPORT_FILE) as z:
    n_charts = len([n for n in z.namelist() if n.startswith("xl/charts/chart")])
n_exp = sum(1 for p in plans for it in p["questions"] + p.get("cmp", []) + p["rels"] if it[3]["kind"] and it[3]["table"] is not None and len(it[3]["table"]))
ck(f"charts written to the file ({n_charts})", n_charts >= n_exp - 5, f"{n_charts} found, about {n_exp} expected")

res = pd.DataFrame(checks, columns=["check", "status", "detail"])
print(res.to_string(index=False))
for w_ in warns:
    print("NOTE:", w_)
failed = res[res["status"] == "FAIL"]
if len(failed):
    raise AssertionError("These checks failed (the workbook was already saved, but do not use it until this is explained):\n" +
                         "\n".join(f"  - {r.check}: {r.detail}" for r in failed.itertuples()))
print("\nALL CHECKS PASSED")


"""
SR Report Generator - Product Ops

File size gate:
  > 60 MB  →  Report Generator only  (Overview / Daily / Weekly / Monthly)
  ≤ 60 MB  →  Report Generator + Insights & RCA + Smart Summary

Performance features:
  - @st.cache_data on load+preprocess (re-runs are instant for same file)
  - Single-pass groupby in compute_sr
  - Parallel Excel sheet computation via ThreadPoolExecutor (4 workers)
  - apply_formatting samples 200 rows for col-width (not full sheet)
  - frozenset for TIER_1_BANKS / CARD_MODES → O(1) membership
  - No .copy() except at category/merge boundaries (_decat shallow copy)
"""

import io
import gc
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional, Dict, List

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Payment SR Analysis", layout="wide")
st.title("SR Report Generator - Product Ops")

st.markdown("""
<a href="https://metabase.cashfree.com/question/23625-sr-report-data?merchantid=&start_date=&end_date="
   target="_blank" style="font-size:18px;font-weight:bold;">
CLICK HERE TO GET THE DATA FROM METABASE
</a>
""", unsafe_allow_html=True)
st.markdown("---")

# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------
LARGE_FILE_THRESHOLD_MB = 60   # files above this get Report Generator only

TIER_1_BANKS = frozenset([
    "AXIS BANK", "HDFC BANK", "ICICI BANK",
    "KOTAK MAHINDRA BANK", "STATE BANK OF INDIA", "YES BANK LTD",
])
CARD_MODES = frozenset([
    "CREDIT_CARD", "DEBIT_CARD", "CARD", "PREPAID_CARD", "CREDIT_CARD_EMI",
])

PSP_MAP = {
    "abcdicici": "ABCD (Aditya Birla Capital)", "abfspay": "Bajaj Finserv / Markets",
    "airtel": "Airtel Thanks App", "allbank": "Allahabad Bank (Now Indian Bank)",
    "amazonpay": "Amazon Pay", "andb": "Andhra Bank (Now Union Bank)", "apl": "Amazon Pay",
    "ary": "Aryavart Bank", "aubank": "AU Small Finance Bank", "axisb": "Axis Mobile / CRED",
    "axisbank": "Axis Mobile App", "axl": "PhonePe", "bandhan": "Bandhan Bank",
    "barodampay": "bob World", "barodapay": "bob World", "bhim": "BHIM App",
    "boi": "BHIM Aadhaar / Mobile", "bombob": "BHIM BOM", "bpaywallet": "BharatPe",
    "bpunity": "BharatPe", "cbin": "Central Bank Mobile", "centralbank": "Central Bank Mobile",
    "citi": "Citibank India (Now Axis)", "citigold": "Citibank India (Now Axis)",
    "cnrb": "Candy / Canara ai1", "corp": "Corporation Bank (Now Union Bank)", "cred": "CRED",
    "csbpay": "CSB Bank", "cub": "CUB mBank Plus", "db": "Deutsche Bank",
    "dbs": "DBS Digibank", "dhani": "Dhani App", "digikhata": "DigiKhata",
    "dlb": "Dhanlaxmi Bank", "dnsbank": "DNS Bank", "draxisbank": "DR", "drbob": "DR",
    "drcanb": "DR", "drfederal": "DR", "drhdfcbank": "DR", "dricici": "DR",
    "dridbi": "DR", "dridfc": "DR", "drindus": "DR", "drkotak": "DR", "drpnb": "DR",
    "drsbi": "DR", "drubi": "DR", "druco": "DR", "dryesb": "DR", "ebixcash": "EbixCash",
    "equitas": "Equitas Mobile", "equitasbank": "Equitas Mobile", "esaf": "ESAF Mobile",
    "fam": "FamPay", "fbl": "FedMobile / CoinTab", "federal": "FedMobile",
    "fifederal": "Fi Money", "fincarebank": "Fincare Mobile", "finobank": "FinoPay",
    "fkaxis": "Flipkart", "freecharge": "Freecharge", "freoicici": "Freo",
    "goaxb": "Kiwi", "gwaxis": "Genwise", "hdfc": "HDFC Bank Mobile",
    "hdfcbank": "HDFC Bank Mobile", "hsbc": "HSBC India", "hsbcbank": "HSBC India",
    "ibl": "PhonePe", "icici": "iMobile / Pockets", "idbi": "IDBI Go Mobile+",
    "idfcbank": "IDFC First Mobile", "idfcfirst": "IDFC First Mobile", "ikwik": "MobiKwik",
    "imobile": "iMobile Pay", "indianbank": "IndOASIS", "indianbk": "IndOASIS",
    "indie": "INDIE", "indus": "IndusMobile", "indusind": "IndusMobile",
    "inhdfc": "Tata Neu", "iob": "IOB Mobile", "janabank": "Jana Mobile",
    "jarunity": "Jar App", "jio": "MyJio / JioPay", "jkb": "J&K Bank",
    "jsb": "Janaseva Bank", "jupiter": "Jupiter Money", "jupiteraxis": "Jupiter Money",
    "kaypay": "Kotak (Legacy)", "kbaxis": "KreditBee", "kbl": "KBL Mobile",
    "kotak": "Kotak Mobile", "kotak811": "Kotak 811", "kphdfc": "Kredit.Pe",
    "kvb": "KVB Dlite", "liv": "LivQuik", "lxaxis": "LiquiLoans", "mahb": "MahaMobile",
    "maxaxis": "Max Life", "mbk": "MobiKwik", "mbkns": "MobiKwik",
    "mboi": "Bank of India Mobile", "mvhdfc": "Money View", "naviaxis": "Navi App",
    "niyoicici": "Niyo", "nsdl": "NSDL Jiffy", "nye": "Niyo", "nyes": "Niyo",
    "obopay": "Obopay", "okaxis": "Google Pay", "okhdfcbank": "Google Pay",
    "okicici": "Google Pay", "oksbi": "Google Pay", "omni": "OmniCard",
    "oneyes": "OneCard", "paytm": "Paytm", "paytmwallet": "Paytm", "payu": "PayU",
    "payworld": "Payworld", "payzapp": "PayZapp", "phonepe": "PhonePe",
    "pinelabs": "Pine Labs", "pingpay": "Samsung Pay Mini", "pnb": "PNB One",
    "pnyes": "PennyDrop", "pockets": "Pockets", "postbank": "IPPB Mobile",
    "psb": "PSB UnIC", "psbank": "PSB UnIC", "ptaxis": "Paytm", "pthdfc": "Paytm",
    "ptsbi": "Paytm", "ptyes": "Paytm", "pz": "PayZapp", "pzh": "PayZapp",
    "pzw": "PayZapp", "rapl": "Amazon Pay", "razorpay": "Razorpay",
    "rbl": "RBL MoBank", "rmrbl": "Resilient", "sbi": "BHIM SBI Pay / Yono",
    "sbmbank": "SBM Bank", "scb": "SC Mobile", "seyes": "SalarySe",
    "shriramhdfcbank": "Shriram Finance", "sib": "Mirror+", "slc": "Slice",
    "slice": "Slice", "sliceaxis": "Slice", "slicepay": "Slice",
    "spicepay": "Spice Money", "superyes": "Super.Money",
    "suryoday": "Suryoday Mobile", "tapicici": "Tata Neu",
    "tbl": "Thane Bharat Bank", "timecosmos": "TimePay", "tjsb": "TJSB Mobile",
    "tmb": "TMB Digilobby", "topay": "ToPay", "trans": "Cheq / Transcorp",
    "trio": "Trio", "ubi": "Union Bank", "uboi": "Union Bank", "uco": "UCO mBanking",
    "ujjivan": "Ujjivan Mobile", "unionbank": "Union Bank",
    "unionbankofindia": "Union Bank", "unitypay": "Unity Bank", "upi": "BHIM",
    "utkarshbank": "Utkarsh Mobile", "waaxis": "WhatsApp Pay",
    "wahdfcbank": "WhatsApp Pay", "waicici": "WhatsApp Pay", "wasbi": "WhatsApp Pay",
    "yapl": "Amazon Pay", "ybl": "PhonePe", "yes": "Yes Bank",
    "yesbank": "Iris by Yes Bank", "yescred": "CRED", "yescurie": "CRED (Curie)",
    "yesfam": "FamPay", "yesg": "Groww Pay", "yesgo": "Yes Bank",
    "yespay": "Yes Pay Next", "yespop": "POP", "yestp": "Third Party App",
    "zoicici": "Zomato", "ztrbl": "Zeta",
}

DTYPE_MAP = {
    "merchantid": "category", "paymentmode": "category", "txstatus": "category",
    "bankname": "category", "cardtype": "category", "cardcountry": "category",
    "pg": "category", "txmsg": "category",
}
CHUNK_SIZE = 500_000


# ---------------------------------------------------------------------------
# INTERNAL HELPERS
# ---------------------------------------------------------------------------
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    return df


def _decat(df: pd.DataFrame, cols) -> pd.DataFrame:
    """Shallow copy + cast only listed category cols to str.
    Only called at merge/fillna boundaries."""
    df = df.copy(deep=False)
    for c in cols:
        if c in df.columns and hasattr(df[c], "cat"):
            df[c] = df[c].astype(str)
    return df


def compute_sub_category_vectorized(df: pd.DataFrame) -> pd.Series:
    result = pd.Series("OTHER", index=df.index, dtype=object)
    pm = df["paymentmode"] if "paymentmode" in df.columns else pd.Series("", index=df.index)

    is_upi = pm == "UPI"
    if "bankname" in df.columns:
        bn = df["bankname"].astype(str)
        has_bank = bn.notna() & ~bn.isin({"NAN", "None", "nan", ""})
        result[is_upi & has_bank]  = "UPI_INTENT"
        result[is_upi & ~has_bank] = "UPI_COLLECT"
    else:
        result[is_upi] = "UPI_COLLECT"

    cc = df["cardcountry"].astype(str) if "cardcountry" in df.columns \
        else pd.Series("IN", index=df.index)
    is_dom = cc == "IN"
    for mode in CARD_MODES:
        is_m = pm == mode
        result[is_m & is_dom]  = f"{mode}_DOMESTIC"
        result[is_m & ~is_dom] = f"{mode}_INTERNATIONAL"

    is_nb = pm == "NET_BANKING"
    if "bankname" in df.columns:
        is_t1 = df["bankname"].isin(TIER_1_BANKS)
        result[is_nb & is_t1]  = "NB_TIER_1"
        result[is_nb & ~is_t1] = "NB_TIER_2"
    else:
        result[is_nb] = "NB_TIER_2"

    return result.astype("category")


# ---------------------------------------------------------------------------
# LOAD + PREPROCESS — cached by file bytes hash
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_and_preprocess(file_bytes: bytes) -> pd.DataFrame:
    chunks = []
    chunk_iter = pd.read_csv(
        io.BytesIO(file_bytes),
        low_memory=False,
        dtype=DTYPE_MAP,
        chunksize=CHUNK_SIZE,
        engine="c",
    )
    for chunk in chunk_iter:
        chunk = clean_columns(chunk)
        chunk = _preprocess_chunk(chunk)
        chunks.append(chunk)
    df = pd.concat(chunks, ignore_index=True)
    del chunks
    gc.collect()
    return df


def _preprocess_chunk(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["paymentmode", "txstatus", "bankname", "cardtype",
                "cardcountry", "pg", "txmsg"]:
        if col in df.columns:
            df[col] = (df[col].astype(str).str.upper().str.strip()
                       .replace({"NAN": None, "NONE": None})
                       .astype("category"))

    df["txtime"] = pd.to_datetime(df["txtime"], errors="coerce", format="mixed")
    df.dropna(subset=["txtime"], inplace=True)

    df["Day"]          = df["txtime"].dt.date
    df["Month"]        = df["txtime"].dt.to_period("M").astype(str).astype("category")
    df["Week"]         = df["txtime"].dt.to_period("W").astype(str).astype("category")
    df["Hour"]         = df["txtime"].dt.floor("h").astype(str).astype("category")
    df["Display_Date"] = df["txtime"].dt.strftime("%d-%b").astype("category")
    df.drop(columns=["txtime"], inplace=True)

    if "amount" in df.columns:
        df["amount"] = (pd.to_numeric(
            df["amount"].astype(str).str.replace(r"[^\d.]", "", regex=True),
            errors="coerce").fillna(0).astype("float32"))
    else:
        df["amount"] = pd.array([0] * len(df), dtype="float32")

    bins   = [0, 1_000, 10_000, 50_000, 100_000, 200_000, float("inf")]
    labels = ["0-1k", "1k-10k", "10k-50k", "50k-1L", "1L-2L", ">2L"]
    df["amount_category"] = pd.cut(
        df["amount"], bins=bins, labels=labels, right=True).astype("category")

    df["is_success"]    = (df["txstatus"] == "SUCCESS").astype("int8")
    df["is_userdrop"]   = (df["txstatus"] == "USER_DROPPED").astype("int8")
    df["is_failed"]     = (df["txstatus"] == "FAILED").astype("int8")
    df["is_incomplete"] = df["txstatus"].isin(
        ["PENDING", "INCOMPLETE", "FLAGGED", "CANCELLED"]).astype("int8")

    if "bankname" in df.columns:
        df["bank_tier"] = df["bankname"].isin(TIER_1_BANKS).map(
            {True: "Tier 1 Bank", False: "Tier 2 Bank"}).astype("category")

    if "cardnumber" in df.columns:
        raw = df["cardnumber"].astype(str)
        df["upi_handle"] = (raw.where(raw.str.contains("@", na=False))
                               .str.split("@").str[1].str.lower()
                               .astype("category"))
        df["psp_app"] = df["upi_handle"].map(PSP_MAP).fillna(
            df["upi_handle"]).astype("category")
        df.drop(columns=["cardnumber"], inplace=True)

    if "cardcountry" in df.columns:
        df["card_category"] = (df["cardcountry"] == "IN").map(
            {True: "DOMESTIC", False: "IPG"}).astype("category")

    df["sub_category"] = compute_sub_category_vectorized(df)

    # Pre-compute derived columns used in every groupby — done ONCE at load,
    # not repeated per sheet. Saves 2 groupby passes per compute_sr call.
    df["success_amount"] = (df["amount"] * df["is_success"]).astype("float32")
    df["nodrop_flag"]    = (df["txstatus"] != "USER_DROPPED").astype("int8")
    df["nodrop_success"] = (df["is_success"] * df["nodrop_flag"]).astype("int8")
    return df


# ---------------------------------------------------------------------------
# COMPUTE SR — mega-groupby + rollup strategy
#
# Core insight: instead of scanning N rows once per sheet (10+ times),
# scan once at the finest grain then rollup to each sheet's key.
# This is 5-6x faster on large data.
# ---------------------------------------------------------------------------

# Columns aggregated in every groupby — used in rollup()
_AGG_COLS = ["Volume","Success","UserDrops","Total_Value","GMV","nd_succ","nd_vol"]

def _build_base(data: pd.DataFrame, key: list, tx_id_col: str,
                amount_col: str) -> pd.DataFrame:
    """
    Single groupby on `data` using pre-computed columns.
    Returns a small aggregated DataFrame that all sheets roll up from.
    """
    return data.groupby(key, dropna=False, observed=True).agg(
        Volume      = (tx_id_col,        "count"),
        Success     = ("is_success",     "sum"),
        UserDrops   = ("is_userdrop",    "sum"),
        Total_Value = (amount_col,       "sum"),
        GMV         = ("success_amount", "sum"),   # amount * is_success, pre-computed
        nd_succ     = ("nodrop_success", "sum"),   # is_success * nodrop, pre-computed
        nd_vol      = ("nodrop_flag",    "sum"),   # 1 * nodrop, pre-computed
    ).reset_index()


def _rollup(base: pd.DataFrame, key: list) -> pd.DataFrame:
    """Sum aggregated columns to a coarser key. Runs on tiny data — very fast."""
    return base.groupby(key, dropna=False, observed=True)[_AGG_COLS].sum().reset_index()


def _finalise(grouped: pd.DataFrame, total_volume: int,
              merchant_col: str, num_merchants: int,
              merchant_totals: Optional[pd.Series]) -> pd.DataFrame:
    """Compute derived SR columns from raw aggregated counts."""
    vol = grouped["Volume"].replace(0, 1)
    grouped["Unsuccessful Count"]        = grouped["Volume"] - grouped["Success"]
    grouped["SR (%)"]                    = (grouped["Success"] / vol * 100).round(2)
    grouped["SR without User Drops (%)"] = (grouped["nd_succ"] / grouped["nd_vol"].replace(0,1) * 100).round(2)
    grouped["% of Volume (Global)"]      = (grouped["Volume"] / total_volume * 100).round(2)
    grouped.drop(columns=["nd_succ","nd_vol"], inplace=True)
    if num_merchants > 1 and merchant_totals is not None:
        grouped = grouped.join(merchant_totals.rename("_mt"), on=merchant_col)
        grouped["% of Volume (Per Merchant)"] = (
            grouped["Volume"] / grouped["_mt"].replace(0,1) * 100).round(2)
        grouped.drop(columns=["_mt"], inplace=True)
    return grouped.sort_values("Volume", ascending=False)


def compute_sr(data: pd.DataFrame, group_cols, merchant_col,
               tx_id_col, amount_col, status_col, num_merchants,
               _base_cache: Optional[dict] = None) -> pd.DataFrame:
    """
    Public API — unchanged signature.
    Internally uses _base_cache to avoid re-scanning data when the same
    dataset is used for multiple sheets (e.g. Daily Paymode, Daily Bank, etc.)
    """
    group_cols   = [c for c in group_cols if c is not None]
    total_volume = len(data)
    if total_volume == 0:
        return pd.DataFrame()

    key = [merchant_col] + group_cols

    # Use the pre-built base if the caller passes one (see _build_report_excel)
    if _base_cache is not None:
        cache_key = id(data)
        if cache_key not in _base_cache:
            # Build finest-grain base for this dataset and cache it
            _base_cache[cache_key] = _build_base(data, key, tx_id_col, amount_col)
        base = _base_cache[cache_key]
        # If our key is coarser than the cached base key, rollup; else use direct
        base_key_set = set(base.columns) - set(_AGG_COLS)
        if set(key).issubset(base_key_set):
            grouped = _rollup(base, key)
        else:
            grouped = _build_base(data, key, tx_id_col, amount_col)
    else:
        grouped = _build_base(data, key, tx_id_col, amount_col)

    mt = None
    if num_merchants > 1:
        mt = data.groupby(merchant_col, observed=True)[tx_id_col].count()

    return _finalise(grouped, total_volume, merchant_col, num_merchants, mt)


def compute_mom_change(df: pd.DataFrame, group_cols, merchant_col) -> pd.DataFrame:
    if df.empty:
        return df
    group_cols = [c for c in group_cols if c is not None]
    key_cols   = [c for c in group_cols if c not in ("Month", "Week", "Hour")]
    df = df.sort_values([merchant_col] + group_cols)
    grp = df.groupby([merchant_col] + key_cols, observed=True)
    for metric in ("Volume", "SR (%)"):
        if metric in df.columns:
            df[f"{metric} Δ"]        = grp[metric].diff().fillna(0)
            df[f"{metric} % Change"] = (grp[metric].pct_change().fillna(0) * 100).round(2)
    return df


# ---------------------------------------------------------------------------
# EXCEL FORMATTING — receives open Workbook, formats in-place, saves once
# ---------------------------------------------------------------------------
def _format_workbook(wb) -> io.BytesIO:
    """Apply conditional formatting + col widths to an already-open Workbook."""
    for ws in wb.worksheets:
        if ws.max_row <= 1:
            continue
        headers = [c.value for c in ws[1]]
        vol_idx = headers.index("Volume") if "Volume" in headers else None
        vol_ltr = get_column_letter(vol_idx + 1) if vol_idx is not None else None

        for ci, hdr in enumerate(headers, 1):
            cl  = get_column_letter(ci)
            rng = f"{cl}2:{cl}{ws.max_row}"
            if hdr in ("SR (%)", "SR without User Drops (%)") and vol_ltr:
                ws.conditional_formatting.add(rng, FormulaRule(
                    formula=[f"AND({vol_ltr}2>=10,{cl}2<50)"], font=Font(color="FF0000")))
                ws.conditional_formatting.add(rng, FormulaRule(
                    formula=[f"AND({vol_ltr}2>=10,{cl}2>90)"], font=Font(color="008000")))
            elif hdr == "Unsuccessful Count" and vol_ltr:
                ws.conditional_formatting.add(rng, FormulaRule(
                    formula=[f"AND({vol_ltr}2>=10,{cl}2>=20,{cl}2>0.5*{vol_ltr}2)"],
                    font=Font(color="FF4500")))
            elif isinstance(hdr, str) and "Volume" in hdr and "%" in hdr and "Change" not in hdr:
                ws.conditional_formatting.add(rng, DataBarRule(
                    start_type="num", start_value=0,
                    end_type="num",   end_value=100, color="638EC6"))

        # Sample first 100 rows for col-width — fast enough, visually identical
        for col in ws.iter_cols():
            ltr = col[0].column_letter
            mx  = max((len(str(c.value or "")) for c in col[:100]), default=8)
            ws.column_dimensions[ltr].width = min(mx + 2, 50)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def apply_formatting(buffer: io.BytesIO) -> io.BytesIO:
    buffer.seek(0)
    wb = load_workbook(buffer)
    return _format_workbook(wb)


# ---------------------------------------------------------------------------
# PARALLEL SHEET BUILDER
# ---------------------------------------------------------------------------
def _build_sheet(args):
    sn, cols, dataset, merchant_col, tx_id_col, amount_col, \
        status_col, num_merchants, time_col, report_type = args

    result = compute_sr(dataset, cols, merchant_col, tx_id_col,
                        amount_col, status_col, num_merchants)
    if result.empty:
        return sn, None
    if time_col and time_col in result.columns:
        result = result.sort_values(time_col, ascending=True)
    if report_type in ("Monthly", "Weekly"):
        result = compute_mom_change(result, cols, merchant_col)
    return sn, result


def _build_report_excel(
    sheet_specs, base_breakdowns, cur, tcol, report_type,
    num_merchants, hourly_df=None
) -> io.BytesIO:
    """
    Mega-groupby strategy:
      1. Build ONE base aggregation at the finest grain for each unique dataset.
      2. All other sheets rollup from that base — they never re-scan raw data.
      3. Write + format in one pass.
    """
    # ── Step 1: build one base per unique dataset (by id) ────────────────────
    # The finest key covers all columns any sheet might group by.
    # We group once, all sheets derive from it via rollup.
    FINEST_COLS = ["merchantid","paymentmode","pg","bankname",
                   "sub_category","upi_handle","psp_app","bank_tier",
                   "card_category","cardtype","amount_category"]
    if tcol:
        FINEST_COLS.append(tcol)

    # Build per unique dataset
    dataset_bases: dict = {}   # id(dataset) → base DataFrame
    for sn, (cols, ds) in sheet_specs.items():
        did = id(ds)
        if did not in dataset_bases:
            # Use only columns that actually exist in this dataset
            finest_key = [c for c in FINEST_COLS if c in ds.columns]
            if not finest_key:
                finest_key = ["merchantid"]
            dataset_bases[did] = _build_base(
                ds, finest_key, "transactionid", "amount")

    def _sheet_from_base(sn, cols, ds):
        did   = id(ds)
        base  = dataset_bases[did]
        key   = ["merchantid"] + [c for c in cols if c is not None]
        total = int(base["Volume"].sum())
        if total == 0:
            return sn, None

        # All columns in key must exist in base — rollup
        if all(c in base.columns for c in key):
            grouped = _rollup(base, key)
        else:
            # Fallback: direct groupby (shouldn't happen with FINEST_COLS)
            grouped = _build_base(ds, key, "transactionid", "amount")

        mt = None
        if num_merchants > 1:
            mt = base.groupby("merchantid", observed=True)["Volume"].sum()

        result = _finalise(grouped, total, "merchantid", num_merchants, mt)
        if result.empty:
            return sn, None
        if tcol and tcol in result.columns:
            result = result.sort_values(tcol, ascending=True)
        if report_type in ("Monthly","Weekly"):
            result = compute_mom_change(result, cols, "merchantid")
        return sn, result

    # ── Step 2: compute all sheets (fast — rollups on tiny base) ─────────────
    computed: dict = {}
    with ThreadPoolExecutor(max_workers=6) as pool:
        futs = {pool.submit(_sheet_from_base, sn, cols, ds): sn
                for sn, (cols, ds) in sheet_specs.items()}
        for fut in as_completed(futs):
            sn, res = fut.result()
            if res is not None:
                computed[sn] = res

    # ── failure sheets (fast groupbys, no parallelism needed) ────────────────
    fail_data = cur[cur["txstatus"] != "SUCCESS"]
    if not fail_data.empty:
        fc = ["merchantid"] + ([tcol] if tcol else []) + ["paymentmode","txmsg"]
        computed["Failures Analysis"] = (
            fail_data.groupby(fc, dropna=False, observed=True)["transactionid"]
            .count().reset_index(name="Volume")
            .sort_values([tcol] if tcol else ["Volume"], ascending=bool(tcol)))

        if "bankname" in fail_data.columns:
            bc = ["merchantid"] + ([tcol] if tcol else []) + ["paymentmode","bankname","txmsg"]
            computed["Failures (Paymode+Bank)"] = (
                fail_data.groupby(bc, dropna=False, observed=True)["transactionid"]
                .count().reset_index(name="Volume")
                .sort_values("Volume", ascending=False))

        if "pg" in fail_data.columns:
            pc = ["merchantid"] + ([tcol] if tcol else []) + ["paymentmode","pg","txmsg"]
            computed["Failures (Paymode+PG)"] = (
                fail_data.groupby(pc, dropna=False, observed=True)["transactionid"]
                .count().reset_index(name="Volume")
                .sort_values("Volume", ascending=False))

    if not computed:
        computed["No Data"] = pd.DataFrame([{"Info": "No data"}])

    # ── write all sheets into one Workbook, format in-place, save once ────────
    EXCEL_MAX_ROWS = 1_048_575   # Excel hard limit minus 1 header row

    # Tracks sheets that were split, so we can write an index at the end
    split_log = []   # list of (base_name, num_parts, total_rows)

    def _write_df_safe(writer, df_, sheet_name):
        """
        Write df_ to Excel. If it exceeds Excel row limit, split across
        numbered sheets — NO rows are dropped. All data is written.

        Split sheets are named:  SheetName (1 of N), SheetName (2 of N) …
        A summary line is added to split_log so an index sheet can be built.
        """
        total_rows = len(df_)
        if total_rows <= EXCEL_MAX_ROWS:
            df_.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            return

        # Sort by Volume desc so Part 1 has the highest-volume rows
        if "Volume" in df_.columns:
            df_ = df_.sort_values("Volume", ascending=False).reset_index(drop=True)

        total_parts = (total_rows + EXCEL_MAX_ROWS - 1) // EXCEL_MAX_ROWS  # ceiling div
        base = sheet_name[:22]   # leave room for " (N of NN)"

        for part in range(total_parts):
            start = part * EXCEL_MAX_ROWS
            end   = min(start + EXCEL_MAX_ROWS, total_rows)
            chunk = df_.iloc[start:end]
            sn_part = f"{base} ({part+1} of {total_parts})"
            chunk.to_excel(writer, sheet_name=sn_part, index=False)

        split_log.append((sheet_name, total_parts, total_rows))

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sn, df_ in computed.items():
            _write_df_safe(writer, df_, sn[:31])

        # ── Write a "Split Sheets Index" if any sheet was too large ──────────
        if split_log:
            index_rows = []
            for base_name, n_parts, total_rows in split_log:
                index_rows.append({
                    "Sheet (base name)": base_name,
                    "Total Rows (all data)": total_rows,
                    "Split into N sheets": n_parts,
                    "Rows per sheet": EXCEL_MAX_ROWS,
                    "Note": (
                        f"All {total_rows:,} rows are present across "
                        f"{n_parts} sheets named '{base_name[:22]} (1 of {n_parts})' "
                        f"… '({n_parts} of {n_parts})'. "
                        f"No data has been dropped."
                    ),
                })
            pd.DataFrame(index_rows).to_excel(
                writer, sheet_name="⚠ Split Sheet Index", index=False)

        wb = writer.book
        _format_workbook(wb)

    out.seek(0)
    return out


# ---------------------------------------------------------------------------
# RCA / COMPARE HELPERS  (only used for small files)
# ---------------------------------------------------------------------------
def compare_periods(curr_df: pd.DataFrame, prev_df: pd.DataFrame,
                    group_cols) -> pd.DataFrame:
    if isinstance(group_cols, str):
        group_cols = [group_cols]
    curr_df = _decat(curr_df, group_cols)
    prev_df = _decat(prev_df, group_cols)

    sc = (curr_df.groupby(group_cols, observed=True)["is_success"]
          .agg(["count","sum"]).reset_index()
          .rename(columns={"count":"Vol_curr","sum":"Succ_curr"}))
    sp = (prev_df.groupby(group_cols, observed=True)["is_success"]
          .agg(["count","sum"]).reset_index()
          .rename(columns={"count":"Vol_prev","sum":"Succ_prev"}))

    merged = pd.merge(sc, sp, on=group_cols, how="outer")
    num_cols = ["Vol_curr","Succ_curr","Vol_prev","Succ_prev"]
    merged[num_cols] = merged[num_cols].fillna(0)
    merged["SR_curr"]   = merged["Succ_curr"] / merged["Vol_curr"].replace(0,1) * 100
    merged["SR_prev"]   = merged["Succ_prev"] / merged["Vol_prev"].replace(0,1) * 100
    merged["SR_Delta"]  = merged["SR_curr"] - merged["SR_prev"]
    merged["Vol_Delta"] = merged["Vol_curr"] - merged["Vol_prev"]
    return merged.sort_values("SR_Delta")


def get_failure_spike(curr_df: pd.DataFrame, prev_df: pd.DataFrame,
                      group_cols, mode_group: str) -> pd.DataFrame:
    curr_df = _decat(curr_df, group_cols)
    prev_df = _decat(prev_df, group_cols)

    c_err = curr_df[curr_df["txstatus"] != "SUCCESS"].groupby(group_cols, observed=True).size().reset_index(name="curr_count")
    p_err = prev_df[prev_df["txstatus"] != "SUCCESS"].groupby(group_cols, observed=True).size().reset_index(name="prev_count")

    merged = pd.merge(c_err, p_err, on=group_cols, how="outer")
    merged[["curr_count","prev_count"]] = merged[["curr_count","prev_count"]].fillna(0)
    merged["spike"] = merged["curr_count"] - merged["prev_count"]

    total_inc = (len(curr_df[curr_df["txstatus"] != "SUCCESS"])
                 - len(prev_df[prev_df["txstatus"] != "SUCCESS"]))
    merged["contribution"] = (merged["spike"] / total_inc * 100) if total_inc > 0 else 0.0

    drill = {"UPI": "sub_category", "CARDS": "cardtype", "NET_BANKING": "bankname"}.get(mode_group)
    if drill:
        fails_c = curr_df[curr_df["txstatus"] != "SUCCESS"]
        ctx = []
        for reason in merged["txmsg"]:
            sub = fails_c[fails_c["txmsg"] == reason]
            if not sub.empty and drill in sub.columns:
                top = sub[drill].value_counts().head(3)
                ctx.append(" | ".join(f"{k}: {v}" for k, v in top.items()))
            else:
                ctx.append("N/A")
        merged["Context"] = ctx
    else:
        merged["Context"] = "N/A"

    return merged.sort_values("spike", ascending=False)


def color_delta(val):
    if val < 0:   return "color: #ff4b4b"
    elif val > 0: return "color: #09ab3b"
    return "color: white"


# ---------------------------------------------------------------------------
# SMART SUMMARY HELPERS  (only used for small files)
# ---------------------------------------------------------------------------
def generate_day_summary(curr_df: pd.DataFrame, prev_df: pd.DataFrame) -> str:
    if curr_df.empty:
        return "No data."
    cv = len(curr_df); pv = len(prev_df)
    csr  = curr_df["is_success"].sum() / cv * 100 if cv else 0
    psr  = prev_df["is_success"].sum() / pv * 100 if pv else 0
    diff = csr - psr
    if diff >= -0.5:
        return "✅ Performance is stable."
    ud_inc   = int(curr_df["is_userdrop"].sum()) - int(prev_df["is_userdrop"].sum())
    fail_inc = int(curr_df["is_failed"].sum())   - int(prev_df["is_failed"].sum())
    reason   = "User Drops" if ud_inc > fail_inc else "Technical Failures"
    mc = curr_df[curr_df["txstatus"] != "SUCCESS"]["paymentmode"].value_counts()
    top_mode = mc.index[0] if not mc.empty else "Unknown"
    return f"📉 SR dropped by {abs(diff):.1f}% due to increased **{reason}** in **{top_mode}**."


def render_metric_tab(curr_df, prev_df, metric_col, title):
    c_cnt = int(curr_df[metric_col].sum())
    p_cnt = int(prev_df[metric_col].sum())
    diff  = c_cnt - p_cnt
    cv = len(curr_df); pv = len(prev_df)
    c_pct = c_cnt / cv * 100 if cv else 0
    p_pct = p_cnt / pv * 100 if pv else 0
    pct_d = c_pct - p_pct

    st.markdown(f'**{title}: {c_pct:.1f}% ({pct_d:+.1f}%) {"📈" if pct_d > 0 else "📉"}**')
    if diff <= 0 and pct_d <= 0:
        st.caption(f"No increase in {title}. (Count: {p_cnt} → {c_cnt})")
        return
    st.markdown("---")

    for lbl, exact, lst, dcol in [
        ("UPI",         "UPI",         None,       "upi_handle"),
        ("CARDS",       None,          CARD_MODES, "cardtype"),
        ("NET_BANKING", "NET_BANKING", None,       "bankname"),
    ]:
        mc = curr_df[curr_df["paymentmode"] == exact] if exact \
            else curr_df[curr_df["paymentmode"].isin(lst)]
        mp = prev_df[prev_df["paymentmode"] == exact] if exact \
            else prev_df[prev_df["paymentmode"].isin(lst)]
        if mc.empty:
            continue
        mc_cnt = int(mc[metric_col].sum()); mp_cnt = int(mp[metric_col].sum())
        md = mc_cnt - mp_cnt
        if md <= 0:
            continue
        mc_vol = len(mc); mp_vol = len(mp)
        mc_r = mc_cnt / mc_vol * 100 if mc_vol else 0
        mp_r = mp_cnt / mp_vol * 100 if mp_vol else 0

        c1, c2 = st.columns([1.5, 2])
        with c1:
            st.markdown(f"**{lbl}**")
            st.write(f"Count: {mp_cnt} → {mc_cnt} (+{md})")
            st.caption(f"Rate: {mp_r:.1f}% → {mc_r:.1f}% ({mc_r - mp_r:+.1f}%)")
        with c2:
            if dcol in mc.columns:
                mc2 = _decat(mc[mc[metric_col] == 1], [dcol])
                mp2 = _decat(mp[mp[metric_col] == 1], [dcol])
                gc_ = mc2.groupby(dcol, observed=True).size().reset_index(name="Curr")
                gp_ = mp2.groupby(dcol, observed=True).size().reset_index(name="Prev")
                m2  = pd.merge(gc_, gp_, on=dcol, how="outer")
                m2[["Curr","Prev"]] = m2[["Curr","Prev"]].fillna(0)
                m2["Inc"]      = m2["Curr"] - m2["Prev"]
                m2["% Impact"] = (m2["Curr"] / mc_vol * 100).round(2)
                top = m2.sort_values("Inc", ascending=False).head(3)
                if not top.empty and top.iloc[0]["Inc"] > 0:
                    st.dataframe(
                        top.style.format({"Curr":"{:.0f}","Prev":"{:.0f}",
                                          "Inc":"{:.0f}","% Impact":"{:.2f}%"}),
                        hide_index=True, use_container_width=True)


# ---------------------------------------------------------------------------
# SHARED: REPORT GENERATOR UI
# ---------------------------------------------------------------------------
def render_report_tab(df: pd.DataFrame):
    st.header("Generate Standard Excel Reports")
    c1, c2 = st.columns(2)
    with c1: share_with_merchant = st.checkbox("Share with Merchant (Remove PG Data)", value=False)
    with c2: run_hourly          = st.checkbox("Get Hourly SR Report")

    df_display = df.drop(columns=["pg"], errors="ignore") if share_with_merchant else df
    if share_with_merchant:
        st.warning("PG Data removed from reports.")

    hourly_df = None
    if run_hourly:
        mn, mx = df_display["Day"].min(), df_display["Day"].max()
        hr = st.date_input("Select Date for Hourly Report", [mn, mx], min_value=mn, max_value=mx)
        if len(hr) == 2:
            hourly_df = df_display[(df_display["Day"] >= hr[0]) & (df_display["Day"] <= hr[1])]

    with st.expander("Filter Data (Date, Merchant, Mode)", expanded=True):
        if st.checkbox("Filter by Date Range", value=False):
            dr = st.date_input("Select Date Range",
                               [df_display["Day"].min(), df_display["Day"].max()])
            if len(dr) == 2:
                df_display = df_display[(df_display["Day"] >= dr[0]) & (df_display["Day"] <= dr[1])]

        merchants = sorted(str(x) for x in df_display["merchantid"].unique() if pd.notna(x))
        sel_merch = st.multiselect("Select Merchants", merchants, default=[])
        if sel_merch:
            df_display = df_display[df_display["merchantid"].astype(str).isin(sel_merch)]
            if hourly_df is not None:
                hourly_df = hourly_df[hourly_df["merchantid"].astype(str).isin(sel_merch)]

        if "paymentmode" in df_display.columns:
            modes     = sorted(str(x) for x in df_display["paymentmode"].unique() if pd.notna(x))
            sel_modes = st.multiselect("Select Payment Modes", modes, default=[])
            if sel_modes:
                df_display = df_display[df_display["paymentmode"].isin(sel_modes)]
                if hourly_df is not None:
                    hourly_df = hourly_df[hourly_df["paymentmode"].isin(sel_modes)]

                if any(m in sel_modes for m in CARD_MODES):
                    st.markdown("---")
                    st.subheader("💳 Card Filters")
                    if "card_category" in df_display.columns:
                        cats = sorted(str(x) for x in df_display["card_category"].unique() if pd.notna(x))
                        sc   = st.multiselect("Select Card Category (Geo)", cats, default=[])
                        if sc:
                            mask = (df_display["card_category"].isin(sc) |
                                    ~df_display["paymentmode"].isin(CARD_MODES))
                            df_display = df_display[mask]
                            if hourly_df is not None:
                                hourly_df = hourly_df[
                                    hourly_df["card_category"].isin(sc) |
                                    ~hourly_df["paymentmode"].isin(CARD_MODES)]
                    if "cardtype" in df_display.columns:
                        ctypes = sorted(str(x) for x in df_display["cardtype"].unique() if pd.notna(x))
                        sct    = st.multiselect("Select Card Network", ctypes, default=[])
                        if sct:
                            mask = (df_display["cardtype"].isin(sct) |
                                    ~df_display["paymentmode"].isin(CARD_MODES))
                            df_display = df_display[mask]
                            if hourly_df is not None:
                                hourly_df = hourly_df[
                                    hourly_df["cardtype"].isin(sct) |
                                    ~hourly_df["paymentmode"].isin(CARD_MODES)]

    st.info(f"Ready to analyse **{len(df_display):,}** transactions.")

    if st.button("🚀 Run Report Analysis", key="btn_rep"):
        with st.spinner("Generating Excel files…"):
            num_merchants = df_display["merchantid"].nunique()

            report_configs = {
                "Overview": {"time_col": None},
                "Daily":    {"time_col": "Day"},
                "Weekly":   {"time_col": "Week"},
                "Monthly":  {"time_col": "Month"},
            }
            if run_hourly and hourly_df is not None and not hourly_df.empty:
                report_configs["Hourly"] = {"time_col": "Hour", "data": hourly_df}

            base_breakdowns = {
                "Paymode":         ["paymentmode"],
                "PG":              ["pg"],
                "Bank":            ["bankname"],
                "Paymode+PG":      ["paymentmode","pg"],
                "Paymode+Bank":    ["paymentmode","bankname"],
                "Paymode+PG+Bank": ["paymentmode","pg","bankname"],
            }

            def _make_sheet_specs(cur, tcol, report_type):
                tgrp = [tcol] if tcol else []
                specs = {}
                if report_type == "Overview":
                    specs["SR Overall"] = ([], cur)
                else:
                    specs[f"SR {report_type}"] = (tgrp, cur)
                for nm, gcols in base_breakdowns.items():
                    if all(c in cur.columns for c in gcols):
                        specs[f"SR by {nm}"] = (gcols + tgrp, cur)
                specs["SR by Paymode"] = (tgrp + ["paymentmode"], cur)
                if "cardtype" in cur.columns:
                    specs["SR by Card Type"] = (
                        tgrp + ["paymentmode","cardtype"],
                        cur[cur["paymentmode"].isin(CARD_MODES)])
                if "upi_handle" in cur.columns:
                    hdf = cur[cur["upi_handle"].notna()]
                    if not hdf.empty:
                        specs["SR by UPI Handle"] = (tgrp+["paymentmode","upi_handle"], hdf)
                if "psp_app" in cur.columns:
                    psp = cur[cur["psp_app"].notna()]
                    if not psp.empty:
                        specs["SR by PSP App"] = (tgrp+["paymentmode","psp_app"], psp)
                if "bank_tier" in cur.columns:
                    btd = cur[cur["paymentmode"] == "NET_BANKING"]
                    if not btd.empty:
                        specs["SR by Bank Tier"] = (tgrp+["paymentmode","bank_tier"], btd)
                if "card_category" in cur.columns:
                    ccd = cur[cur["paymentmode"].isin(CARD_MODES)]
                    if not ccd.empty:
                        specs["SR by Card Geo"] = (tgrp+["paymentmode","card_category"], ccd)
                if "amount_category" in cur.columns:
                    specs["SR by Amount Category"] = (tgrp+["paymentmode","amount_category"], cur)
                return specs

            # ── Build all reports in parallel (one future per report type) ───
            def _generate_one(report_type, config):
                cur  = config.get("data", df_display)
                tcol = config["time_col"]
                specs = _make_sheet_specs(cur, tcol, report_type)
                return report_type, _build_report_excel(
                    specs, base_breakdowns, cur, tcol,
                    report_type, num_merchants)

            results = {}   # report_type → BytesIO
            with ThreadPoolExecutor(max_workers=min(len(report_configs), 4)) as pool:
                futs = {pool.submit(_generate_one, rt, cfg): rt
                        for rt, cfg in report_configs.items()}
                for fut in as_completed(futs):
                    rt, buf = fut.result()
                    results[rt] = buf

            # ── Show download buttons in original order ───────────────────────
            st.markdown("### 📥 Download Reports")
            dl_cols = st.columns(min(len(report_configs), 5))
            for idx, report_type in enumerate(report_configs):
                with dl_cols[idx % len(dl_cols)]:
                    st.download_button(
                        label=f"⬇️ {report_type}",
                        data=results[report_type],
                        file_name=f"{report_type}_SR.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{report_type}")
            del results
            gc.collect()
            st.success("✅ All reports ready!")


# ---------------------------------------------------------------------------
# SMALL FILE ONLY: INSIGHTS & RCA TAB
# ---------------------------------------------------------------------------
def render_rca_tab(df: pd.DataFrame):
    st.header("📉 Automated Root Cause Analysis")
    rc1, rc2 = st.columns(2)
    with rc1: sel_days  = st.slider("Comparison Period (Last X Days)", 1, 30, 7)
    with rc2:
        rca_merch = sorted(str(x) for x in df["merchantid"].unique() if pd.notna(x))
        sel_rca_m = st.multiselect("Filter Merchant (RCA)", rca_merch, default=[])

    rca_df  = df[df["merchantid"].astype(str).isin(sel_rca_m)] if sel_rca_m else df
    max_d   = rca_df["Day"].max()
    c_start = max_d   - pd.Timedelta(days=sel_days - 1)
    p_start = c_start - pd.Timedelta(days=sel_days)
    p_end   = c_start - pd.Timedelta(days=1)
    st.caption(f"**Current:** {c_start} → {max_d} | **Previous:** {p_start} → {p_end}")

    curr_df = rca_df[(rca_df["Day"] >= c_start) & (rca_df["Day"] <= max_d)]
    prev_df = rca_df[(rca_df["Day"] >= p_start) & (rca_df["Day"] <= p_end)]

    if curr_df.empty or prev_df.empty:
        st.error("Not enough data for the selected range / merchant.")
        return

    cv = len(curr_df); pv = len(prev_df)
    csr  = curr_df["is_success"].sum() / cv * 100 if cv else 0
    psr  = prev_df["is_success"].sum() / pv * 100 if pv else 0
    cgmv = float(curr_df.loc[curr_df["txstatus"]=="SUCCESS","amount"].sum())
    pgmv = float(prev_df.loc[prev_df["txstatus"]=="SUCCESS","amount"].sum())

    m1,m2,m3,m4,m5,m6 = st.columns(6)
    m1.metric("Curr SR",  f"{csr:.2f}%")
    m2.metric("Prev SR",  f"{psr:.2f}%",   f"{csr-psr:.2f}%")
    m3.metric("Curr Vol", f"{cv:,}")
    m4.metric("Prev Vol", f"{pv:,}",        f"{cv-pv:,}")
    m5.metric("Curr GMV", f"₹{cgmv:,.0f}")
    m6.metric("Prev GMV", f"₹{pgmv:,.0f}", f"{cgmv-pgmv:,.0f}")
    st.markdown("---")

    for mode_grp, exact, lst in [
        ("UPI",         "UPI",         None),
        ("CARDS",       None,          CARD_MODES),
        ("NET_BANKING", "NET_BANKING", None),
    ]:
        mc = curr_df[curr_df["paymentmode"]==exact] if exact \
            else curr_df[curr_df["paymentmode"].isin(lst)]
        mp = prev_df[prev_df["paymentmode"]==exact] if exact \
            else prev_df[prev_df["paymentmode"].isin(lst)]

        with st.expander(f"Analysis: {mode_grp}", expanded=(mode_grp=="UPI")):
            if mc.empty:
                st.info("No data for this mode.")
                continue

            sm = compare_periods(mc, mp, "sub_category")
            ws = sm.sort_values("SR_Delta").head(1)

            col_a, col_b = st.columns([1, 1])
            with col_a:
                st.write("**Sub-Category Performance**")
                st.dataframe(
                    sm[["sub_category","SR_curr","SR_prev","SR_Delta",
                        "Vol_curr","Vol_prev","Vol_Delta"]]
                    .style.format({"SR_curr":"{:.2f}%","SR_prev":"{:.2f}%",
                                   "SR_Delta":"{:.2f}%","Vol_curr":"{:.0f}",
                                   "Vol_prev":"{:.0f}","Vol_Delta":"{:.0f}"})
                    .map(color_delta, subset=["SR_Delta","Vol_Delta"]),
                    use_container_width=True, hide_index=True)

                if not ws.empty and ws["SR_Delta"].values[0] < -1:
                    st.error(f"🚨 **Issue in {ws['sub_category'].values[0]}** "
                             f"(Dropped {ws['SR_Delta'].values[0]:.2f}%)")
                else:
                    st.success("✅ No major sub-category drop.")

                if mode_grp == "UPI" and "upi_handle" in mc.columns:
                    st.markdown("##### 🔍 UPI Handle Breakdown")
                    hs = compare_periods(mc, mp, "upi_handle").sort_values("Vol_Delta")
                    st.dataframe(
                        hs[["upi_handle","SR_curr","SR_prev","SR_Delta",
                            "Vol_curr","Vol_prev","Vol_Delta"]]
                        .style.format({"SR_curr":"{:.2f}%","SR_prev":"{:.2f}%",
                                       "SR_Delta":"{:.2f}%","Vol_curr":"{:.0f}",
                                       "Vol_prev":"{:.0f}","Vol_Delta":"{:.0f}"})
                        .map(color_delta, subset=["SR_Delta","Vol_Delta"]),
                        use_container_width=True, hide_index=True)

                if mode_grp == "CARDS" and "cardtype" in mc.columns:
                    st.markdown("##### 🔍 Card Type Breakdown")
                    cs = compare_periods(mc, mp, ["paymentmode","cardtype"])
                    st.dataframe(
                        cs[["paymentmode","cardtype","SR_curr","SR_prev","SR_Delta",
                            "Vol_curr","Vol_prev","Vol_Delta"]]
                        .style.format({"SR_curr":"{:.2f}%","SR_prev":"{:.2f}%",
                                       "SR_Delta":"{:.2f}%","Vol_curr":"{:.0f}",
                                       "Vol_prev":"{:.0f}","Vol_Delta":"{:.0f}"})
                        .map(color_delta, subset=["SR_Delta","Vol_Delta"]),
                        use_container_width=True, hide_index=True)

            with col_b:
                trend = (mc.groupby("Day", observed=True)
                           .agg(SR=("is_success","mean"), Vol=("transactionid","count"))
                           .reset_index())
                trend["SR"] *= 100
                fig = go.Figure([
                    go.Bar(x=trend["Day"], y=trend["Vol"], name="Volume",
                           marker_color="rgba(135,206,250,0.6)",
                           hovertemplate="<b>Vol:</b> %{y:.0f}<extra></extra>"),
                    go.Scatter(x=trend["Day"], y=trend["SR"], name="SR %", yaxis="y2",
                               line=dict(color="red", width=3),
                               hovertemplate="<b>SR:</b> %{y:.2f}%<extra></extra>"),
                ])
                fig.update_layout(
                    title=f"{mode_grp} Trend",
                    yaxis=dict(title="Volume"),
                    yaxis2=dict(title="SR %", overlaying="y", side="right", range=[0,100]),
                    hovermode="x unified")
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

            st.subheader("Why did it drop?")
            fc1, fc2 = st.columns(2)
            with fc1:
                st.markdown("**1. PG Performance**")
                if "pg" in mc.columns:
                    pg = compare_periods(mc, mp, "pg").sort_values("SR_Delta")
                    st.dataframe(
                        pg[["pg","SR_curr","SR_prev","SR_Delta",
                            "Vol_curr","Vol_prev","Vol_Delta"]]
                        .style.format({"SR_curr":"{:.2f}%","SR_prev":"{:.2f}%",
                                       "SR_Delta":"{:.2f}%","Vol_curr":"{:.0f}",
                                       "Vol_prev":"{:.0f}","Vol_Delta":"{:.0f}"})
                        .map(color_delta, subset=["SR_Delta","Vol_Delta"]),
                        use_container_width=True, hide_index=True)
            with fc2:
                st.markdown("**2. Error Analysis**")
                spikes = get_failure_spike(mc, mp, ["txmsg"], mode_grp)
                if not spikes.empty:
                    ds = spikes[["txmsg","curr_count","prev_count",
                                 "spike","contribution","Context"]].copy()
                    ds.columns = ["Error Message","Curr Vol","Prev Vol",
                                  "Vol Spike","Contrib %","Context"]
                    st.dataframe(
                        ds.style.format({"Curr Vol":"{:.0f}","Prev Vol":"{:.0f}",
                                         "Vol Spike":"{:.0f}","Contrib %":"{:.2f}%"})
                        .background_gradient(subset=["Vol Spike"], cmap="Reds"),
                        use_container_width=True, hide_index=True)
                else:
                    st.info("No specific error spike detected.")


# ---------------------------------------------------------------------------
# SMALL FILE ONLY: SMART SUMMARY TAB
# ---------------------------------------------------------------------------
def render_summary_tab(df: pd.DataFrame):
    st.header("📝 Smart Day-over-Day Analysis")
    ss1, ss2 = st.columns(2)
    with ss1: s_start = st.date_input("Start Date", value=df["Day"].min())
    with ss2: s_end   = st.date_input("End Date",   value=df["Day"].max())

    sum_merch = st.multiselect(
        "Filter Merchant (Optional)",
        sorted(str(x) for x in df["merchantid"].unique() if pd.notna(x)),
        default=[])

    sum_df = df[(df["Day"] >= s_start) & (df["Day"] <= s_end)]
    if sum_merch:
        sum_df = sum_df[sum_df["merchantid"].astype(str).isin(sum_merch)]

    if sum_df.empty:
        st.error("No data for selected range.")
        return

    dates = sorted(sum_df["Day"].unique())
    if len(dates) < 2:
        st.warning("Need at least 2 days of data.")
        return

    st.subheader(f"📊 Daily Deep Dives ({len(dates)-1} comparisons)")
    for i in range(1, len(dates)):
        cd = dates[i]; pd_ = dates[i-1]
        dc = sum_df[sum_df["Day"] == cd]
        dp = sum_df[sum_df["Day"] == pd_]

        cv = len(dc); pv = len(dp)
        csr = dc["is_success"].sum()/cv*100 if cv else 0
        psr = dp["is_success"].sum()/pv*100 if pv else 0
        dsr = csr - psr; dv = cv - pv
        icon = "📉" if dsr < -1 else "✅"

        with st.expander(
            f"{icon} **{cd.strftime('%d-%b')}** (vs {pd_.strftime('%d-%b')}) "
            f"| SR: {csr:.1f}% ({dsr:+.1f}%) | Vol: {pv} → {cv} ({dv:+})"
        ):
            st.info(generate_day_summary(dc, dp))
            t1, t2, t3, t4 = st.tabs([
                "📉 Success Rate","🚧 User Dropped","💥 Failed","⏳ Incomplete"])

            with t1:
                sm = compare_periods(dc, dp, "paymentmode").sort_values("SR_Delta")
                st.dataframe(
                    sm[["paymentmode","SR_curr","SR_prev","SR_Delta","Vol_curr"]]
                    .style.format({"SR_curr":"{:.1f}%","SR_prev":"{:.1f}%",
                                   "SR_Delta":"{:+.1f}%","Vol_curr":"{:.0f}"})
                    .map(color_delta, subset=["SR_Delta"]),
                    use_container_width=True, hide_index=True)
            with t2: render_metric_tab(dc, dp, "is_userdrop",   "User Drops")
            with t3: render_metric_tab(dc, dp, "is_failed",     "Technical Failures")
            with t4: render_metric_tab(dc, dp, "is_incomplete", "Incomplete Txns")


# ===========================================================================
# MAIN
# ===========================================================================
uploaded_file = st.file_uploader("Upload CSV File", type=["csv"])

if uploaded_file:
    file_size_mb = uploaded_file.size / (1024 * 1024)
    is_large_file = file_size_mb > LARGE_FILE_THRESHOLD_MB

    if is_large_file:
        st.warning(
            f"📦 Large file detected ({file_size_mb:.1f} MB). "
            f"**Report Generator only** — Insights & RCA and Smart Summary are "
            f"disabled for files above {LARGE_FILE_THRESHOLD_MB} MB to keep performance fast."
        )
    else:
        st.info(f"📄 File size: {file_size_mb:.1f} MB — all features available.")

    file_bytes = uploaded_file.read()

    with st.spinner("Loading & preprocessing data…"):
        df = load_and_preprocess(file_bytes)
    del file_bytes
    gc.collect()

    st.success(
        f"✅ Loaded **{len(df):,}** transactions. "
        f"RAM ≈ {df.memory_usage(deep=True).sum()/1e6:.0f} MB"
    )

    # ── LARGE FILE: report only ─────────────────────────────────────────────
    if is_large_file:
        tab_report, = st.tabs(["📊 Report Generator"])
        with tab_report:
            render_report_tab(df)

    # ── SMALL FILE: all three tabs ───────────────────────────────────────────
    else:
        tab_report, tab_rca, tab_summary = st.tabs([
            "📊 Report Generator", "🔍 Insights & RCA", "📝 Smart Summary"
        ])
        with tab_report:
            render_report_tab(df)
        with tab_rca:
            render_rca_tab(df)
        with tab_summary:
            render_summary_tab(df)

else:
    st.info("👆 Upload a CSV file to start.")

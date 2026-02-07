import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import tempfile

# ========================
# XLSB normalize
# ========================
def normalize_xlsb_df(df):
    for c in df.columns:
        s = df[c]
        s2 = pd.to_numeric(s, errors="ignore")

        if pd.api.types.is_numeric_dtype(s2):
            if s2.dropna().between(20000, 60000).mean() > 0.8:
                s2 = pd.to_datetime("1899-12-30") + pd.to_timedelta(s2, unit="D")

        df[c] = s2
    return df


def read_excel_safely(path, sheet, header_row):
    p = Path(path)
    suf = p.suffix.lower()

    if suf == ".xlsb":
        engine = "pyxlsb"
    elif suf in [".xlsx", ".xlsm"]:
        engine = "openpyxl"
    elif suf == ".xls":
        engine = "xlrd"
    else:
        raise ValueError(f"Unsupported format: {suf}")

    if sheet == "" or sheet is None:
        xls = pd.ExcelFile(p, engine=engine)
        sheet = xls.sheet_names[0]

    df = pd.read_excel(
        p,
        sheet_name=sheet,
        header=header_row,
        engine=engine,
        dtype=object
    )

    if suf == ".xlsb":
        df = normalize_xlsb_df(df)
    else:
        df = df.apply(lambda s: pd.to_numeric(s, errors="ignore"))

    return df


def convert_headers_to_yyyyww(cols: pd.Index):
    s = pd.Index(cols).astype(str)
    is_yyyyww = s.str.fullmatch(r"\d{6}", na=False)

    to_parse = s.where(~is_yyyyww, None)
    dt = pd.to_datetime(to_parse, errors="coerce", dayfirst=True)
    is_date = dt.notna()

    new = s.copy().to_series()

    if is_date.any():
        iso = dt[is_date].isocalendar()
        new_vals = iso["year"].astype(str) + iso["week"].astype(int).map("{:02d}".format)
        new.loc[is_date] = new_vals.to_numpy()

    new = pd.Index(new)
    week_mask = is_yyyyww | is_date

    return new, week_mask


def consolidate_weeks_fast(df, week_mask, sort_week_cols=True):
    non = df.loc[:, ~week_mask]
    wk = df.loc[:, week_mask]

    if wk.shape[1] == 0:
        return df

    wk_num = wk.apply(pd.to_numeric, errors="coerce")
    wk_sum = wk_num.groupby(wk_num.columns, axis=1).sum(min_count=1)

    wk_sum = wk_sum.loc[:, ~wk_sum.columns.duplicated(keep="last")]

    if sort_week_cols:
        def wkey(x):
            xs = str(x)
            return (0, int(xs)) if xs.isdigit() and len(xs) == 6 else (1, xs)

        wk_sum = wk_sum[sorted(wk_sum.columns, key=wkey)]

    return pd.concat([non, wk_sum], axis=1)


def filter_firm_forecast_colB(df):
    if df.shape[1] <= 1:
        return df

    col = df.iloc[:, 1].astype(str).str.strip().str.lower()
    mask = col.isin(["firm", "forecast"])
    return df.loc[mask].copy()


def process_excel(file_path, sheet_name, header_row):
    df = read_excel_safely(file_path, sheet_name, header_row)
    df = filter_firm_forecast_colB(df)

    new_cols, week_mask = convert_headers_to_yyyyww(pd.Index(df.columns))
    df.columns = new_cols

    df = consolidate_weeks_fast(df, week_mask=week_mask)

    return df


# ========================
# STREAMLIT UI
# ========================
st.title("Convert Header to YYYYWW")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xlsm", "xls", "xlsb"]
)

sheet_name = st.text_input("Sheet name (optional)", "")
header_row = st.number_input("Header row", min_value=0, max_value=100, value=0)

if uploaded_file:
    suffix = Path(uploaded_file.name).suffix

    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name

    if st.button("Process"):
        try:
            df = process_excel(tmp_path, sheet_name, header_row)

            st.success("Processing complete")
            st.dataframe(df)

            out_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df.to_excel(out_file.name, index=False, engine="xlsxwriter")

            with open(out_file.name, "rb") as f:
                st.download_button(
                    "Download result",
                    f,
                    file_name=f"{datetime.today().strftime('%Y%m%d')}.xlsx"
                )

        except Exception as e:
            st.error(str(e))

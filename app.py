import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import tempfile

# ========================
# Helpers
# ========================
def read_excel_sheets(path):
    p = Path(path)
    suf = p.suffix.lower()

    if suf == ".xlsb":
        engine = "pyxlsb"
    elif suf in [".xlsx", ".xlsm"]:
        engine = "openpyxl"
    elif suf == ".xls":
        engine = "xlrd"
    else:
        raise ValueError("Unsupported format")

    xls = pd.ExcelFile(p, engine=engine)
    return xls, xls.sheet_names, engine


def read_sheet(path, sheet, header_row, engine):
    df = pd.read_excel(
        path,
        sheet_name=sheet,
        header=header_row,
        engine=engine,
        dtype=object
    )
    return df


# ========================
# Processing (giữ nguyên logic)
# ========================
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


def consolidate_weeks_fast(df, week_mask):
    non = df.loc[:, ~week_mask]
    wk = df.loc[:, week_mask]

    if wk.shape[1] == 0:
        return df

    wk_num = wk.apply(pd.to_numeric, errors="coerce")
    wk_sum = wk_num.groupby(wk_num.columns, axis=1).sum(min_count=1)
    wk_sum = wk_sum.loc[:, ~wk_sum.columns.duplicated(keep="last")]

    return pd.concat([non, wk_sum], axis=1)


def filter_firm_forecast_colB(df):
    if df.shape[1] <= 1:
        return df

    col = df.iloc[:, 1].astype(str).str.strip().str.lower()
    mask = col.isin(["firm", "forecast"])
    return df.loc[mask].copy()


# ========================
# STREAMLIT UI
# ========================
st.title("Convert Header to YYYYWW-Tin Vo")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xlsm", "xls", "xlsb"]
)

if uploaded_file:
    suffix = Path(uploaded_file.name).suffix

    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name

    try:
        xls, sheet_names, engine = read_excel_sheets(tmp_path)

        sheet_selected = st.selectbox("Select sheet", sheet_names)
        header_row = st.number_input("Header row", 0, 50, 0)

        preview = pd.read_excel(
            tmp_path,
            sheet_name=sheet_selected,
            header=None,
            engine=engine,
            nrows=10
        )

        st.write("Preview (raw):")
        st.dataframe(preview)

        if st.button("Process"):
            df = read_sheet(tmp_path, sheet_selected, header_row, engine)
            df = filter_firm_forecast_colB(df)

            new_cols, week_mask = convert_headers_to_yyyyww(pd.Index(df.columns))
            df.columns = new_cols
            df = consolidate_weeks_fast(df, week_mask)

            st.success("Done")
            st.dataframe(df)

            out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df.to_excel(out.name, index=False)

            with open(out.name, "rb") as f:
                st.download_button(
                    "Download result",
                    f,
                    file_name=f"{datetime.today().strftime('%Y%m%d')}.xlsx"
                )

    except Exception as e:
        st.error(str(e))

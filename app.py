import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

st.set_page_config(page_title="PDF to Excel AI", page_icon="📄")

st.title("📄 PDF → Excel AI")
st.write("PDF upload karo, command do aur clean Excel pao.")


def unique_columns(columns):
    result = []
    used = {}

    for col in columns:
        col = str(col).strip()

        if not col:
            col = "Column"

        if col in used:
            used[col] += 1
            result.append(f"{col}_{used[col]}")
        else:
            used[col] = 0
            result.append(col)

    return result


def extract_pdf(pdf_file):
    tables = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            try:
                page_tables = page.extract_tables()
            except Exception:
                page_tables = []

            for table in page_tables:
                if table and len(table) > 1:
                    tables.append(table)

    if not tables:
        return None

    table = tables[0]

    header = table[0]
    rows = table[1:]

    header = [
        str(x).strip() if x is not None else ""
        for x in header
    ]

    header = [
        x if x else f"Column_{i+1}"
        for i, x in enumerate(header)
    ]

    header = unique_columns(header)

    clean_rows = []

    for row in rows:
        row = list(row)

        if len(row) < len(header):
            row += [""] * (len(header) - len(row))

        if len(row) > len(header):
            row = row[:len(header)]

        row = [
            str(x).strip() if x is not None else ""
            for x in row
        ]

        if any(x != "" for x in row):
            clean_rows.append(row)

    if not clean_rows:
        return None

    return pd.DataFrame(
        clean_rows,
        columns=header
    )


def clean_spaces(df):
    df = df.copy()

    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: re.sub(
                r"\s+",
                " ",
                str(x)
            ).strip()
        )

    return df


def process_command(df, command):
    result = df.copy()
    cmd = command.lower().strip()

    # Duplicate rows
    if any(x in cmd for x in [
        "duplicate",
        "duplicates",
        "डुप्लीकेट",
        "डुप्लिकेट",
        "दोहराव"
    ]):
        result = result.drop_duplicates()

    # Blank rows
    if any(x in cmd for x in [
        "blank",
        "empty",
        "खाली",
        "blank row",
        "empty row"
    ]):
        result = result.replace(
            r"^\s*$",
            pd.NA,
            regex=True
        )
        result = result.dropna(
            how="all"
        )

    # Extra spaces
    if any(x in cmd for x in [
        "space",
        "spaces",
        "trim",
        "स्पेस",
        "खाली जगह"
    ]):
        result = clean_spaces(result)

    # Column names
    if any(x in cmd for x in [
        "column",
        "columns",
        "column name",
        "column names",
        "कॉलम"
    ]):
        new_cols = []

        for col in result.columns:
            col = str(col).strip()
            col = re.sub(r"\s+", "_", col)
            new_cols.append(col)

        result.columns = unique_columns(
            new_cols
        )

    # A-Z
    if any(x in cmd for x in [
        "a-z",
        "a to z",
        "a se z",
        "ascending",
        "ए से जेड"
    ]):
        if len(result.columns) > 0:
            result = result.sort_values(
                by=result.columns[0],
                ascending=True
            )

    # Z-A
    if any(x in cmd for x in [
        "z-a",
        "z to a",
        "z se a",
        "descending",
        "जेड से ए"
    ]):
        if len(result.columns) > 0:
            result = result.sort_values(
                by=result.columns[0],
                ascending=False
            )

    result.columns = unique_columns(
        result.columns
    )

    return result.reset_index(drop=True)


def make_excel(df):
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Clean Data"
        )

    output.seek(0)
    return output


uploaded = st.file_uploader(
    "📄 PDF upload karo",
    type=["pdf"]
)

if uploaded:

    if st.button("🔍 PDF ka Data Read Karo"):

        try:
            df = extract_pdf(uploaded)

            if df is None:
                st.error(
                    "PDF me table data nahi mila."
                )
            else:
                st.session_state["df"] = df
                st.success(
                    "PDF data successfully read ho gaya ✅"
                )

        except Exception as e:
            st.error(
                f"PDF error: {e}"
            )


if "df" in st.session_state:

    st.subheader("📋 Original Data")

    st.dataframe(
        st.session_state["df"],
        use_container_width=True
    )

    command = st.text_input(
        "🤖 Command do",
        placeholder=(
            "duplicate hatao"
        )
    )

    if st.button("⚡ Clean Data"):

        if not command.strip():
            st.warning(
                "Pehle command likho."
            )
        else:
            try:
                cleaned = process_command(
                    st.session_state["df"],
                    command
                )

                st.session_state[
                    "cleaned"
                ] = cleaned

                st.success(
                    "Data clean ho gaya ✅"
                )

            except Exception as e:
                st.error(
                    f"Processing error: {e}"
                )


if "cleaned" in st.session_state:

    st.subheader("✨ Cleaned Data")

    st.dataframe(
        st.session_state["cleaned"],
        use_container_width=True
    )

    excel = make_excel(
        st.session_state["cleaned"]
    )

    st.download_button(
        "📥 Excel Download Karo",
        data=excel,
        file_name="cleaned_data.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(
    page_title="PDF to Excel AI",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF → Excel AI")
st.write("PDF upload karo, Hindi/English command do, aur clean Excel file pao.")


# =========================
# PDF DATA EXTRACT
# =========================
def extract_pdf_data(pdf_file):
    all_tables = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                if table and len(table) > 1:
                    all_tables.append(table)

    if not all_tables:
        return None

    header = all_tables[0]
    rows = all_tables[0][1:]

    header = [
        str(x).strip()
        if x is not None and str(x).strip()
        else f"Column_{i+1}"
        for i, x in enumerate(header)
    ]

    clean_rows = []

    for row in rows:
        if not row:
            continue

        row = [
            str(x).strip() if x is not None else ""
            for x in row
        ]

        if len(row) < len(header):
            row = row + [""] * (len(header) - len(row))

        if len(row) > len(header):
            row = row[:len(header)]

        clean_rows.append(row)

    if not clean_rows:
        return None

    return pd.DataFrame(clean_rows, columns=header)


# =========================
# AI COMMAND ENGINE
# =========================
def process_command(df, command):

    command = command.lower().strip()
    result = df.copy()

    # -------------------------
    # 1. EMPTY / BLANK ROW
    # -------------------------
    empty_words = [
        "empty",
        "blank",
        "empty row",
        "blank row",
        "खाली",
        "खाली row",
        "खाली लाइन",
        "खाली पंक्ति"
    ]

    if any(word in command for word in empty_words):

        result = result.dropna(how="all")

        for col in result.columns:
            result[col] = result[col].replace(
                r"^\s*$",
                pd.NA,
                regex=True
            )

        result = result.dropna(how="all")


    # -------------------------
    # 2. DUPLICATE
    # -------------------------
    duplicate_words = [
        "duplicate",
        "duplicates",
        "duplication",
        "डुप्लीकेट",
        "डुप्लिकेट",
        "दोहराया",
        "दोहराव"
    ]

    if any(word in command for word in duplicate_words):

        result = result.drop_duplicates()


    # -------------------------
    # 3. EXTRA SPACE
    # -------------------------
    space_words = [
        "extra space",
        "extra spaces",
        "spaces remove",
        "space remove",
        "trim",
        "स्पेस",
        "extra खाली जगह"
    ]

    if any(word in command for word in space_words):

        for col in result.columns:

            if result[col].dtype == "object":

                result[col] = (
                    result[col]
                    .astype(str)
                    .str.replace(
                        r"\s+",
                        " ",
                        regex=True
                    )
                    .str.strip()
                )


    # -------------------------
    # 4. MISSING VALUES
    # -------------------------
    missing_words = [
        "missing",
        "null",
        "n/a",
        "missing value",
        "खाली value",
        "missing data",
        "null हटाओ",
        "na हटाओ"
    ]

    if any(word in command for word in missing_words):

        result = result.replace(
            [
                "",
                "nan",
                "NaN",
                "None",
                "none",
                "NULL",
                "null",
                "NA",
                "N/A"
            ],
            pd.NA
        )


    # -------------------------
    # 5. COLUMN CLEAN
    # -------------------------
    column_words = [
        "column clean",
        "clean column",
        "column name",
        "rename column",
        "columns clean",
        "column के नाम",
        "column ka naam",
        "column name clean"
    ]

    if any(word in command for word in column_words):

        new_columns = []

        for col in result.columns:

            col = str(col).strip()

            col = re.sub(
                r"\s+",
                "_",
                col
            )

            col = re.sub(
                r"[^a-zA-Z0-9_अ-ह]+",
                "",
                col
            )

            new_columns.append(
                col.lower()
            )

        result.columns = new_columns


    # -------------------------
    # 6. NUMBER CLEAN
    # -------------------------
    number_words = [
        "number clean",
        "numeric",
        "numbers clean",
        "संख्या साफ",
        "number ko clean"
    ]

    if any(word in command for word in number_words):

        for col in result.columns:

            converted = pd.to_numeric(
                result[col]
                .astype(str)
                .str.replace(
                    ",",
                    "",
                    regex=False
                ),
                errors="coerce"
            )

            if converted.notna().sum() > 0:
                result[col] = converted


    # -------------------------
    # 7. A-Z SORT
    # -------------------------
    sort_words = [
        "sort",
        "a-z",
        "a to z",
        "ascending",
        "क्रम",
        "क्रम में",
        "ए से जेड",
        "a से z"
    ]

    if any(word in command for word in sort_words):

        sort_col = None

        # Command me column ka naam check karo
        for col in result.columns:

            if str(col).lower() in command:
                sort_col = col
                break

        # Name column automatically find
        if sort_col is None:

            for col in result.columns:

                if any(
                    x in str(col).lower()
                    for x in [
                        "name",
                        "naam",
                        "नाम"
                    ]
                ):
                    sort_col = col
                    break

        # Agar column nahi mila
        # to first column
        if sort_col is None and len(result.columns) > 0:
            sort_col = result.columns[0]

        if sort_col is not None:

            result = result.sort_values(
                by=sort_col,
                ascending=True,
                na_position="last"
            )


    # -------------------------
    # 8. Z-A SORT
    # -------------------------
    descending_words = [
        "z-a",
        "z to a",
        "descending",
        "उल्टा क्रम",
        "बड़े से छोटे"
    ]

    if any(
        word in command
        for word in descending_words
    ):

        sort_col = None

        for col in result.columns:

            if str(col).lower() in command:
                sort_col = col
                break

        if sort_col is None:

            for col in result.columns:

                if any(
                    x in str(col).lower()
                    for x in [
                        "name",
                        "naam",
                        "नाम"
                    ]
                ):
                    sort_col = col
                    break

        if sort_col is None and len(result.columns) > 0:
            sort_col = result.columns[0]

        if sort_col is not None:

            result = result.sort_values(
                by=sort_col,
                ascending=False,
                na_position="last"
            )


    # -------------------------
    # FINAL SPACE CLEAN
    # -------------------------
    for col in result.columns:

        if result[col].dtype == "object":

            result[col] = result[col].map(
                lambda x:
                x.strip()
                if isinstance(x, str)
                else x
            )

    return result


# =========================
# CREATE EXCEL
# =========================
def create_excel(df):

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


# =========================
# PDF UPLOAD
# =========================
uploaded_file = st.file_uploader(
    "📄 PDF upload karo",
    type=["pdf"]
)


if uploaded_file:

    st.success(
        "PDF upload ho gaya ✅"
    )

    if st.button(
        "🔍 PDF ka Data Read Karo"
    ):

        with st.spinner(
            "PDF read ho raha hai..."
        ):

            try:

                df = extract_pdf_data(
                    uploaded_file
                )

                if df is None:

                    st.error(
                        "PDF me table data nahi mila. "
                        "Text-based/table PDF try karo."
                    )

                else:

                    st.session_state["df"] = df

                    st.session_state.pop(
                        "cleaned_df",
                        None
                    )

                    st.success(
                        f"{len(df)} rows aur "
                        f"{len(df.columns)} columns mile."
                    )

            except Exception as e:

                st.error(
                    f"PDF read error: {e}"
                )


# =========================
# ORIGINAL DATA
# =========================
if "df" in st.session_state:

    st.subheader(
        "📋 Original Data"
    )

    st.dataframe(
        st.session_state["df"],
        use_container_width=True
    )


    # =====================
    # COMMAND
    # =====================
    st.subheader(
        "🤖 Command do"
    )

    st.info(
        "Example: duplicate hatao, "
        "blank rows hatao aur naam A-Z me lagao"
    )

    command = st.text_input(
        "Apna command yahan likho...",
        placeholder=(
            "Jaise: duplicate hatao "
            "aur extra spaces remove karo"
        )
    )


    if st.button(
        "⚡ Clean Data"
    ):

        if not command.strip():

            st.warning(
                "Pehle command likho."
            )

        else:

            with st.spinner(
                "Data clean ho raha hai..."
            ):

                try:

                    cleaned_df = process_command(
                        st.session_state["df"].copy(),
                        command
                    )

                    st.session_state[
                        "cleaned_df"
                    ] = cleaned_df

                    st.success(
                        "Data clean ho gaya ✅"
                    )

                except Exception as e:

                    st.error(
                        f"Processing error: {e}"
                    )


# =========================
# CLEANED DATA
# =========================
if "cleaned_df" in st.session_state:

    st.subheader(
        "✨ Cleaned Data"
    )

    st.dataframe(
        st.session_state["cleaned_df"],
        use_container_width=True
    )


    # Excel file
    excel_file = create_excel(
        st.session_state["cleaned_df"]
    )


    st.download_button(
        label="📥 Excel Download Karo",
        data=excel_file,
        file_name="cleaned_data.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

import streamlit as st
import pandas as pd
import pdfplumber
import re
import io


# =========================
# PAGE SETTINGS
# =========================

st.set_page_config(
    page_title="PDF to Excel AI",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF → Excel AI")
st.write(
    "PDF upload karo, Hindi/English command do, "
    "aur clean Excel file pao."
)


# =========================
# MAKE UNIQUE COLUMN NAMES
# =========================

def make_unique_columns(columns):
    """
    Duplicate column names ko automatically unique banata hai.
    Example:
    Name, Name, Name
    →
    Name, Name_1, Name_2
    """

    seen = {}
    unique_columns = []

    for col in columns:

        col = str(col).strip()

        if not col:
            col = "Column"

        if col not in seen:
            seen[col] = 0
            unique_columns.append(col)

        else:
            seen[col] += 1
            unique_columns.append(
                f"{col}_{seen[col]}"
            )

    return unique_columns


# =========================
# PDF DATA EXTRACTION
# =========================

def extract_pdf_data(pdf_file):

    all_tables = []

    try:

        with pdfplumber.open(pdf_file) as pdf:

            for page in pdf.pages:

                try:
                    tables = page.extract_tables()
                except Exception:
                    tables = []

                for table in tables:

                    if table and len(table) >= 2:
                        all_tables.append(table)

    except Exception as e:

        raise ValueError(
            f"PDF read nahi ho paya: {e}"
        )


    if not all_tables:
        return None


    # First table use karo
    table = all_tables[0]

    header = table[0]
    rows = table[1:]


    # =========================
    # CLEAN HEADER
    # =========================

    clean_header = []

    for i, value in enumerate(header):

        if value is None:
            value = ""

        value = str(value).strip()

        if not value:
            value = f"Column_{i + 1}"

        clean_header.append(value)


    # IMPORTANT:
    # Duplicate column names fix
    clean_header = make_unique_columns(
        clean_header
    )


    # =========================
    # CLEAN ROWS
    # =========================

    clean_rows = []

    for row in rows:

        if row is None:
            continue

        row = list(row)

        # Short row
        if len(row) < len(clean_header):

            row = row + (
                [""] *
                (len(clean_header) - len(row))
            )

        # Long row
        elif len(row) > len(clean_header):

            row = row[:len(clean_header)]


        new_row = []

        for value in row:

            if value is None:
                value = ""

            value = str(value).strip()

            new_row.append(value)


        # Empty row check
        if any(
            str(x).strip()
            for x in new_row
        ):
            clean_rows.append(new_row)


    if not clean_rows:
        return None


    df = pd.DataFrame(
        clean_rows,
        columns=clean_header
    )


    # Final safety check
    df.columns = make_unique_columns(
        df.columns
    )

    return df


# =========================
# TEXT CLEAN FUNCTION
# =========================

def clean_text_columns(df):

    result = df.copy()

    for col in result.columns:

        if (
            pd.api.types.is_object_dtype(
                result[col]
            )
        ):

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

    return result


# =========================
# REMOVE BLANK ROWS
# =========================

def remove_blank_rows(df):

    result = df.copy()

    result = result.replace(
        r"^\s*$",
        pd.NA,
        regex=True
    )

    result = result.dropna(
        how="all"
    )

    return result.reset_index(
        drop=True
    )


# =========================
# REMOVE DUPLICATES
# =========================

def remove_duplicates(df):

    result = df.copy()

    result = result.drop_duplicates()

    return result.reset_index(
        drop=True
    )


# =========================
# CLEAN MISSING VALUES
# =========================

def clean_missing_values(df):

    result = df.copy()

    missing_values = [
        "",
        " ",
        "nan",
        "NaN",
        "NAN",
        "None",
        "none",
        "NULL",
        "null",
        "NA",
        "N/A",
        "n/a",
        "-"
    ]

    result = result.replace(
        missing_values,
        pd.NA
    )

    return result


# =========================
# CLEAN COLUMN NAMES
# =========================

def clean_column_names(df):

    result = df.copy()

    new_columns = []

    for i, col in enumerate(
        result.columns
    ):

        col = str(col).strip()

        # Space → underscore
        col = re.sub(
            r"\s+",
            "_",
            col
        )

        # Special characters remove
        col = re.sub(
            r"[^a-zA-Z0-9_\u0900-\u097F]+",
            "",
            col
        )

        if not col:
            col = f"Column_{i + 1}"

        new_columns.append(
            col.lower()
        )

    # Duplicate names fix
    result.columns = make_unique_columns(
        new_columns
    )

    return result


# =========================
# CLEAN NUMBERS
# =========================

def clean_numbers(df):

    result = df.copy()

    for col in result.columns:

        original = result[col]

        converted = pd.to_numeric(
            original
            .astype(str)
            .str.replace(
                ",",
                "",
                regex=False
            )
            .str.replace(
                "₹",
                "",
                regex=False
            )
            .str.strip(),
            errors="coerce"
        )

        # Sirf tab convert karo jab
        # reasonable amount numeric ho
        if converted.notna().sum() > 0:

            result[col] = converted.where(
                converted.notna(),
                original
            )

    return result


# =========================
# SORT COLUMN FINDER
# =========================

def find_sort_column(df, command):

    command = command.lower()

    # Command me column name search
    for col in df.columns:

        if str(col).lower() in command:
            return col


    # Common name columns
    name_words = [
        "name",
        "naam",
        "नाम",
        "customer",
        "student",
        "employee",
        "person"
    ]

    for col in df.columns:

        col_text = str(col).lower()

        if any(
            word in col_text
            for word in name_words
        ):
            return col


    # First column fallback
    if len(df.columns) > 0:
        return df.columns[0]


    return None


# =========================
# MAIN COMMAND ENGINE
# =========================

def process_command(df, command):

    result = df.copy()

    command = str(
        command
    ).lower().strip()


    # =========================
    # 1. BLANK / EMPTY ROWS
    # =========================

    blank_words = [
        "blank",
        "empty",
        "blank

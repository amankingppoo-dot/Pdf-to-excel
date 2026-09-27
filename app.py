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
st.write("PDF upload karo, command do, aur clean Excel file pao.")

# PDF se tables/data nikalna
def extract_pdf_data(pdf_file):
    all_tables = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                if table and len(table) > 1:
                    all_tables.extend(table)

    if not all_tables:
        return None

    # First row ko header maan rahe hain
    header = all_tables[0]
    rows = all_tables[1:]

    # Empty values clean
    header = [
        str(x).strip() if x is not None else f"Column_{i+1}"
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

        # Column count same karna
        if len(row) < len(header):
            row += [""] * (len(header) - len(row))

        if len(row) > len(header):
            row = row[:len(header)]

        clean_rows.append(row)

    if not clean_rows:
        return None

    return pd.DataFrame(clean_rows, columns=header)


# Command ke according data clean karna
def process_command(df, command):
    command = command.lower().strip()

    # Empty rows remove
    if "empty" in command or "blank" in command:
        df = df.dropna(how="all")

        for col in df.columns:
            df[col] = df[col].replace(r"^\s*$", pd.NA, regex=True)

        df = df.dropna(how="all")

    # Duplicate remove
    if "duplicate" in command or "duplicates" in command:
        df = df.drop_duplicates()

    # Column names clean
    if "column" in command and ("clean" in command or "rename" in command):
        new_columns = []

        for col in df.columns:
            col = str(col).strip()
            col = re.sub(r"\s+", "_", col)
            col = re.sub(r"[^a-zA-Z0-9_]", "", col)
            new_columns.append(col.lower())

        df.columns = new_columns

    # Extra spaces remove
    if "space" in command or "spaces" in command:
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].str.replace(r"\s+", " ", regex=True)

    # Empty cells ko NA banana
    if "missing" in command or "null" in command:
        df = df.replace(["", "nan", "None", "-"], pd.NA)

    # Rows sort
    if "sort" in command:
        first_col = df.columns[0]
        df = df.sort_values(by=first_col, na_position="last")

    return df


# Excel banana
def create_excel(df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Clean Data"
        )

    output.seek(0)
    return output


# Upload
uploaded_file = st.file_uploader(
    "📤 PDF upload karo",
    type=["pdf"]
)

if uploaded_file:

    st.success("PDF upload ho gaya ✅")

    if st.button("🔍 PDF ka Data Read Karo"):

        with st.spinner("PDF read ho raha hai..."):
            try:
                df = extract_pdf_data(uploaded_file)

                if df is None:
                    st.error(
                        "PDF me table data nahi mila. "
                        "Text-based/table PDF try karo."
                    )
                else:
                    st.session_state["df"] = df

                    st.success(
                        f"{len(df)} rows aur "
                        f"{len(df.columns)} columns mile."
                    )

            except Exception as e:
                st.error(f"PDF read error: {e}")


# Agar data mil gaya
if "df" in st.session_state:

    st.subheader("📊 Original Data")
    st.dataframe(
        st.session_state["df"],
        use_container_width=True
    )

    st.subheader("🤖 Command do")

    command = st.text_input(
        "Example: duplicate hatao, blank rows hatao aur columns clean karo",
        placeholder="Apna command yahan likho..."
    )

    if st.button("⚡ Clean Data"):

        if not command.strip():
            st.warning("Pehle command likho.")
        else:

            with st.spinner("Data clean ho raha hai..."):

                try:
                    cleaned_df = process_command(
                        st.session_state["df"].copy(),
                        command
                    )

                    st.session_state["cleaned_df"] = cleaned_df

                    st.success("Data clean ho gaya ✅")

                except Exception as e:
                    st.error(f"Processing error: {e}")


# Cleaned result
if "cleaned_df" in st.session_state:

    st.subheader("✨ Cleaned Data")

    st.dataframe(
        st.session_state["cleaned_df"],
        use_container_width=True
    )

    excel_file = create_excel(
        st.session_state["cleaned_df"]
    )

    st.download_button(
        label="📥 Excel Download Karo",
        data=excel_file,
        file_name="cleaned_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

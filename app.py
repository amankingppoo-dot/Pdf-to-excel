import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from datetime import datetime

st.set_page_config(
    page_title="PDF → Excel AI",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF → Excel AI")
st.caption("PDF upload karo → natural language command do → Excel ready.")


# =========================================================
# TEXT / COLUMN HELPERS
# =========================================================

def norm(x):
    x = str(x).lower().strip()
    x = x.replace("_", " ").replace("-", " ")
    return re.sub(r"\s+", " ", x)


ALIASES = {
    "amount": [
        "amount", "amt", "राशि", "रकम",
        "कीमत", "मूल्य", "price",
        "total", "total amount", "payment"
    ],

    "name": [
        "name", "naam", "नाम",
        "customer", "customer name"
    ],

    "city": [
        "city", "शहर",
        "स्थान", "location"
    ],

    "status": [
        "status", "स्थिति"
    ],

    "date": [
        "date", "दिनांक",
        "तारीख", "दिन"
    ],

    "phone": [
        "phone", "mobile",
        "मोबाइल", "फोन", "contact"
    ],

    "email": [
        "email", "mail", "ईमेल"
    ],

    "id": [
        "id", "code",
        "क्रमांक", "number", "no"
    ],
}


def find_col(df, command, preferred=None):

    cmd = norm(command)

    if preferred:
        for c in df.columns:
            if norm(c) == norm(preferred):
                return c

    # Exact column name
    for c in df.columns:
        nc = norm(c)

        if nc and nc in cmd:
            return c

    # Alias groups
    for group, aliases in ALIASES.items():

        if any(a in cmd for a in aliases):

            for c in df.columns:

                nc = norm(c)

                if any(
                    a == nc or a in nc
                    for a in aliases
                ):
                    return c

    return None


def numeric_values(s):

    x = (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(
            r"[₹$€£]",
            "",
            regex=True
        )
        .str.extract(
            r"(-?\d+(?:\.\d+)?)",
            expand=False
        )
    )

    return pd.to_numeric(
        x,
        errors="coerce"
    )


def make_excel(df, summary=None):

    out = io.BytesIO()

    with pd.ExcelWriter(
        out,
        engine="openpyxl"
    ):

        df.to_excel(
            out,
            index=False,
            sheet_name="Clean Data"
        )

        if summary is not None and len(summary):

            summary.to_excel(
                out,
                index=False,
                sheet_name="Summary"
            )

    out.seek(0)

    return out


# =========================================================
# PDF EXTRACTION
# =========================================================

def extract_pdf_data(pdf_file):

    tables_all = []

    with pdfplumber.open(pdf_file) as pdf:

        for page_no, page in enumerate(
            pdf.pages,
            1
        ):

            tables = page.extract_tables()

            for table in tables:

                if table and len(table) > 1:

                    tables_all.append(
                        (page_no, table)
                    )

    if not tables_all:
        return None

    first_table = tables_all[0][1]

    header = first_table[0]

    ncols = len(header)

    rows = []

    for _, table in tables_all:

        if not table:
            continue

        local_header = table[0]

        local_rows = table[1:]

        if len(local_header) != ncols:
            continue

        rows.extend(local_rows)

    header = [

        str(v).strip()

        if v is not None
        and str(v).strip()

        else f"Column_{i+1}"

        for i, v in enumerate(header)
    ]

    clean = []

    for row in rows:

        if row is None:
            continue

        row = [
            ""
            if v is None
            else str(v).strip()
            for v in row
        ]

        row = (
            row
            + [""] * ncols
        )[:ncols]

        if any(
            str(v).strip()
            for v in row
        ):

            clean.append(row)

    if not clean:
        return None

    return pd.DataFrame(
        clean,
        columns=header
    )


# =========================================================
# COMMAND HELPERS
# =========================================================

def has_any(cmd, words):

    return any(
        w in cmd
        for w in words
    )


# =========================================================
# SORT
# =========================================================

def command_sort(df, cmd):

    asc_words = [

        "low to high",
        "low se high",

        "kam se zyada",
        "kam se adhik",

        "कम से ज्यादा",
        "कम से अधिक",

        "small to big",
        "smallest first",

        "lowest first",

        "ascending",

        "a-z",
        "a to z",
        "a se z",

        "ए से जेड"
    ]

    desc_words = [

        "high to low",
        "high se low",

        "zyada se kam",
        "jyada se kam",

        "बड़े से छोटे",
        "बड़ा से छोटा",

        "large to small",
        "largest first",

        "highest first",

        "descending",

        "z-a",
        "z to a",
        "z se a",

        "जेड से ए"
    ]

    asc = has_any(
        cmd,
        asc_words
    )

    desc = has_any(
        cmd,
        desc_words
    )

    if (
        "low niche" in cmd
        and
        "high upar" in cmd
    ):

        desc = True
        asc = False

    if not (asc or desc):
        return df, False

    col = find_col(
        df,
        cmd
    )

    amount_requested = has_any(
        cmd,
        ALIASES["amount"]
    )

    if (
        col is None
        and amount_requested
    ):

        for c in df.columns:

            if any(
                a in norm(c)
                for a in ALIASES["amount"]
            ):

                col = c
                break

    # Default common columns
    if col is None:

        for group in [
            "name",
            "date",
            "amount"
        ]:

            for c in df.columns:

                if any(
                    a in norm(c)
                    for a in ALIASES[group]
                ):

                    col = c
                    break

            if col is not None:
                break

    if (
        col is None
        and len(df.columns)
    ):

        col = df.columns[0]

    if col is None:
        return df, False

    numeric = numeric_values(
        df[col]
    )

    numeric_like = (
        numeric.notna().sum()
        >= max(
            1,
            int(len(df) * 0.5)
        )
    )

    if (
        amount_requested
        or any(
            a in norm(col)
            for a in ALIASES["amount"]
        )
        or numeric_like
    ):

        tmp = df.copy()

        tmp["__sort__"] = numeric

        tmp = tmp.sort_values(
            "__sort__",
            ascending=asc,
            na_position="last"
        )

        return (
            tmp.drop(
                columns="__sort__"
            ),
            True
        )

    return (
        df.sort_values(
            col,
            ascending=asc,
            na_position="last",
            key=lambda s:
                s.astype(str).str.lower()
        ),
        True
    )


# =========================================================
# FILTER
# =========================================================

def parse_filter(df, cmd):

    patterns = [

        (
            r"(?:amount|amt|राशि|रकम|price|कीमत)"
            r".{0,20}"
            r"(?:above|over|greater than|more than|से ज्यादा|से अधिक)"
            r"\s*([0-9][0-9,]*)",
            ">"
        ),

        (
            r"(?:amount|amt|राशि|रकम|price|कीमत)"
            r".{0,20}"
            r"(?:below|under|less than|से कम)"
            r"\s*([0-9][0-9,]*)",
            "<"
        ),

        (
            r"(?:amount|amt|राशि|रकम|price|कीमत)"
            r"\s*[><]"
            r"\s*([0-9][0-9,]*)",
            None
        )
    ]

    for pat, op in patterns:

        m = re.search(
            pat,
            cmd
        )

        if m:

            col = find_col(
                df,
                cmd
            )

            if col is None:

                for c in df.columns:

                    if any(
                        a in norm(c)
                        for a in ALIASES["amount"]
                    ):

                        col = c
                        break

            if col:

                val = float(
                    m.group(1)
                    .replace(",", "")
                )

                nums = numeric_values(
                    df[col]
                )

                if op == ">":

                    return (
                        df[nums > val],
                        True
                    )

                if op == "<":

                    return (
                        df[nums < val],
                        True
                    )

                symbol = (
                    ">"
                    if ">" in m.group(0)
                    else "<"
                )

                if symbol == ">":

                    return (
                        df[nums > val],
                        True
                    )

                return (
                    df[nums < val],
                    True
                )

    # City / Status / Name
    for group in [
        "city",
        "status",
        "name"
    ]:

        aliases = ALIASES[group]

        for alias in aliases:

            m = re.search(

                re.escape(alias)
                + r"\s*"
                + r"(?:is|=|:|me|में)?"
                + r"\s*"
                + r"([A-Za-z][A-Za-z0-9 _-]*)",

                cmd
            )

            if m:

                value = m.group(1).strip()

                value = re.split(
                    r"\s+(?:and|aur|or|ya)\s+",
                    value
                )[0].strip()

                for c in df.columns:

                    if any(
                        a in norm(c)
                        for a in aliases
                    ):

                        mask = (
                            df[c]
                            .astype(str)
                            .str.strip()
                            .str.lower()
                            == value.lower()
                        )

                        if mask.any():

                            return (
                                df[mask],
                                True
                            )

    return df, False


# =========================================================
# SUMMARY
# =========================================================

def summary_dataframe(df):

    rows = []

    rows.append(
        [
            "Rows",
            len(df)
        ]
    )

    rows.append(
        [
            "Columns",
            len(df.columns)
        ]
    )

    for c in df.columns:

        nums = numeric_values(
            df[c]
        )

        if (
            nums.notna().sum()
            >= max(
                1,
                int(len(df) * 0.5)
            )
        ):

            rows.append(
                [
                    f"Total {c}",
                    float(nums.sum())
                ]
            )

            rows.append(
                [
                    f"Average {c}",
                    (
                        float(nums.mean())
                        if nums.notna().any()
                        else 0
                    )
                ]
            )

    return pd.DataFrame(
        rows,
        columns=[
            "Metric",
            "Value"
        ]
    )


# =========================================================
# MAIN COMMAND ENGINE
# =========================================================

def process_command(
    df,
    raw_command
):

    cmd = norm(
        raw_command
    )

    result = df.copy()

    actions = []

    # ---------------------------------
    # Blank rows
    # ---------------------------------

    if has_any(
        cmd,
        [
            "empty row",
            "blank row",
            "blank",
            "empty",
            "खाली row",
            "खाली लाइन",
            "खाली पंक्ति",
            "खाली"
        ]
    ):

        result = result.dropna(
            how="all"
        )

        for c in result.columns:

            result[c] = result[c].replace(
                r"^\s*$",
                pd.NA,
                regex=True
            )

        result = result.dropna(
            how="all"
        )

        actions.append(
            "blank rows removed"
        )

    # ---------------------------------
    # Duplicate
    # ---------------------------------

    if has_any(
        cmd,
        [
            "duplicate",
            "duplicates",
            "duplication",
            "डुप्लीकेट",
            "डुप्लिकेट",
            "दोहराव",
            "दोहराया"
        ]
    ):

        before = len(result)

        result = result.drop_duplicates()

        actions.append(
            f"{before-len(result)} duplicate rows removed"
        )

    # ---------------------------------
    # Extra spaces
    # ---------------------------------

    if has_any(
        cmd,
        [
            "extra space",
            "extra spaces",
            "space remove",
            "spaces remove",
            "trim",
            "स्पेस",
            "खाली जगह"
        ]
    ):

        for c in result.columns:

            if result[c].dtype == "object":

                result[c] = (
                    result[c]
                    .astype(str)
                    .str.replace(
                        r"\s+",
                        " ",
                        regex=True
                    )
                    .str.strip()
                )

        actions.append(
            "extra spaces cleaned"
        )

    # ---------------------------------
    # Missing values
    # ---------------------------------

    if has_any(
        cmd,
        [
            "missing",
            "null",
            "n/a",
            "na hatao",
            "missing value",
            "missing data",
            "खाली value"
        ]
    ):

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
                "N/A",
                "n/a"
            ],

            pd.NA
        )

        actions.append(
            "missing values normalized"
        )

    # ---------------------------------
    # Missing rows remove
    # ---------------------------------

    if has_any(
        cmd,
        [
            "missing row hatao",
            "null row hatao",
            "rows with missing",
            "missing rows remove"
        ]
    ):

        result = result.dropna(
            how="any"
        )

        actions.append(
            "rows containing missing values removed"
        )

    # ---------------------------------
    # Clean column names
    # ---------------------------------

    if has_any(
        cmd,
        [
            "column clean",
            "clean column",
            "column names clean",
            "column name clean",
            "columns clean",
            "column ke naam"
        ]
    ):

        new = []

        for c in result.columns:

            x = re.sub(
                r"\s+",
                "_",
                str(c).strip()
            )

            x = re.sub(
                r"[^a-zA-Z0-9_अ-ह]",
                "",
                x
            )

            new.append(
                x.lower()
                or "column"
            )

        result.columns = new

        actions.append(
            "column names cleaned"
        )

    # ---------------------------------
    # Number cleanup
    # ---------------------------------

    if has_any(
        cmd,
        [
            "number clean",
            "numbers clean",
            "numeric clean",
            "number ko clean",
            "संख्या साफ"
        ]
    ):

        for c in result.columns:

            nums = numeric_values(
                result[c]
            )

            if (
                nums.notna().sum()
                >= max(
                    1,
                    int(len(result) * 0.7)
                )
            ):

                result[c] = nums

        actions.append(
            "numeric values cleaned"
        )

    # ---------------------------------
    # Filter
    # ---------------------------------

    result, filtered = parse_filter(
        result,
        cmd
    )

    if filtered:

        actions.append(
            "filter applied"
        )

    # ---------------------------------
    # Sort
    # ---------------------------------

    result, sorted_ok = command_sort(
        result,
        cmd
    )

    if sorted_ok:

        if any(
            w in cmd
            for w in [
                "low to high",
                "low se high",
                "ascending",
                "a-z",
                "a to z",
                "kam se zyada"
            ]
        ):

            direction = "ascending"

        else:

            direction = "descending"

        actions.append(
            f"sorted {direction}"
        )

    # ---------------------------------
    # Selected columns
    # ---------------------------------

    if has_any(
        cmd,
        [
            "only columns",
            "sirf columns",
            "sirf column"
        ]
    ):

        requested = []

        for c in result.columns:

            if norm(c) in cmd:

                requested.append(c)

        if requested:

            result = result[
                requested
            ]

            actions.append(
                "selected columns kept"
            )

    # ---------------------------------
    # Remove column
    # ---------------------------------

    m = re.search(

        r"(?:remove|delete|drop|hatao)"
        r"\s+(?:column|col)"
        r"\s+([a-zA-Z0-9_]+)",

        cmd
    )

    if m:

        target = m.group(1).strip()

        for c in list(
            result.columns
        ):

            if norm(c) == target:

                result = result.drop(
                    columns=[c]
                )

                actions.append(
                    f"column {c} removed"
                )

    # ---------------------------------
    # Serial number
    # ---------------------------------

    if has_any(
        cmd,
        [
            "serial number",
            "sr no",
            "s.no",
            "क्रमांक",
            "row number"
        ]
    ):

        result.insert(
            0,
            "Sr_No",
            range(
                1,
                len(result) + 1
            )
        )

        actions.append(
            "serial number added"
        )

    # ---------------------------------
    # Summary
    # ---------------------------------

    summary_requested = has_any(
        cmd,
        [
            "summary",
            "report",
            "total",
            "sum",
            "average",
            "avg",
            "count",
            "कुल",
            "औसत",
            "रिपोर्ट"
        ]
    )

    summary = (
        summary_dataframe(result)
        if summary_requested
        else None
    )

    if summary_requested:

        actions.append(
            "summary generated"
        )

    # ---------------------------------
    # Final cleanup
    # ---------------------------------

    for c in result.columns:

        if result[c].dtype == "object":

            result[c] = result[c].map(

                lambda x:
                x.strip()
                if isinstance(x, str)
                else x
            )

    return (
        result,
        summary,
        actions
    )


# =========================================================
# SESSION STATE
# =========================================================

if "df" not in st.session_state:
    st.session_state["df"] = None

if "cleaned_df" not in st.session_state:
    st.session_state["cleaned_df"] = None

if "summary_df" not in st.session_state:
    st.session_state["summary_df"] = None

if "history" not in st.session_state:
    st.session_state["history"] = []


# =========================================================
# UPLOAD
# =========================================================

uploaded = st.file_uploader(
    "📄 PDF upload karo",
    type=["pdf"]
)


if uploaded:

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

                data = extract_pdf_data(
                    uploaded
                )

                if data is None:

                    st.error(
                        "PDF me readable table data nahi mila."
                    )

                else:

                    st.session_state[
                        "df"
                    ] = data

                    st.session_state[
                        "cleaned_df"
                    ] = data.copy()

                    st.session_state[
                        "summary_df"
                    ] = None

                    st.session_state[
                        "history"
                    ] = []

                    st.success(

                        f"{len(data)} rows aur "
                        f"{len(data.columns)} columns mile ✅"
                    )

                    st.rerun()

            except Exception as e:

                st.error(
                    f"PDF read error: {e}"
                )


# =========================================================
# MAIN UI
# =========================================================

if st.session_state["df"] is not None:

    current = st.session_state[
        "cleaned_df"
    ]

    st.subheader(
        "📋 Current Excel Data"
    )

    st.dataframe(
        current,
        use_container_width=True,
        height=380
    )

    st.subheader(
        "🤖 AJ Command"
    )

    st.caption(
        "Hindi/English me normal language me command do. "
        "Ek ke baad doosri command bhi chalegi."
    )

    examples = [

        "duplicate hatao",

        "amount low to high karo",

        "amount high to low karo",

        "blank rows hatao",

        "extra spaces hatao",

        "name A-Z karo",

        "amount 1000 se zyada wale dikhao",

        "sirf Raipur ka data rakho",

        "summary bana do",

        "total amount batao",

        "report bana do"
    ]

    st.write(
        "**Examples:** "
        + " • ".join(
            examples[:6]
        )
    )

    # ---------------------------------
    # COMMAND FORM
    # ---------------------------------

    with st.form(
        "command_form",
        clear_on_submit=True
    ):

        command = st.text_input(

            "Command",

            placeholder=
            "Jaise: duplicate hatao aur amount low to high karo"
        )

        submitted = st.form_submit_button(

            "⚡ Command Apply Karo"
        )

    if submitted:

        if not command.strip():

            st.warning(
                "Pehle command likho."
            )

        else:

            with st.spinner(
                "AI-style command process ho rahi hai..."
            ):

                try:

                    (
                        new_df,
                        summary,
                        actions
                    ) = process_command(

                        current,
                        command
                    )

                    st.session_state[
                        "cleaned_df"
                    ] = new_df

                    st.session_state[
                        "summary_df"
                    ] = summary

                    st.session_state[
                        "history"
                    ].append(

                        {
                            "time":
                            datetime.now().strftime(
                                "%H:%M:%S"
                            ),

                            "command":
                            command,

                            "actions":
                            (
                                ", ".join(actions)
                                if actions
                                else
                                "No matching operation"
                            )
                        }
                    )

                    st.rerun()

                except Exception as e:

                    st.error(
                        f"Command error: {e}"
                    )


# =========================================================
# SUMMARY
# =========================================================

if (
    st.session_state[
        "summary_df"
    ] is not None
):

    st.subheader(
        "📊 Summary / Report"
    )

    st.dataframe(

        st.session_state[
            "summary_df"
        ],

        use_container_width=True
    )


# =========================================================
# COMMAND HISTORY
# =========================================================

st.subheader(
    "🧾 Command History"
)

if st.session_state[
    "history"
]:

    st.dataframe(

        pd.DataFrame(
            st.session_state[
                "history"
            ]
        ),

        use_container_width=True,

        hide_index=True
    )

else:

    st.caption(
        "Abhi koi command history nahi hai."
    )


# =========================================================
# DOWNLOAD + RESET
# =========================================================

if st.session_state["df"] is not None:

    col1, col2 = st.columns(2)

    with col1:

        excel = make_excel(

            st.session_state[
                "cleaned_df"
            ],

            st.session_state[
                "summary_df"
            ]
        )

        st.download_button(

            "📥 Final Excel Download Karo",

            data=excel,

            file_name=
            "AI_cleaned_excel.xlsx",

            mime=
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            use_container_width=True
        )

    with col2:

        if st.button(
            "↩️ Original Data Par Reset Karo",
            use_container_width=True
        ):

            st.session_state[
                "cleaned_df"
            ] = st.session_state[
                "df"
            ].copy()

            st.session_state[
                "summary_df"
            ] = None

            st.session_state[
                "history"
            ] = []

            st.rerun()


st.divider()

st.caption(
    "Core processing local Python/Pandas based hai; "
    "paid AI API required nahi hai."
)

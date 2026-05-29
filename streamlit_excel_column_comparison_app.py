import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


MIN_WORD_LENGTH = 4

STOPWORDS = {
    "the", "and", "for", "with", "from", "that", "this",
    "are", "was", "were", "have", "has", "into", "onto",
    "shall", "must", "should", "will", "their", "there",
    "you", "your", "its", "they", "them", "than", "then",
    "also", "such", "each", "when", "where", "using", "based"
}

CONTROL_ID_PATTERN = re.compile(
    r"\bCORE\s*#?\s*(\d+(?:\.\d+)?)\b",
    re.IGNORECASE
)

ISE_ID_PATTERN = re.compile(
    r"^\s*(\d+(?:\.\d+)?)\s*$"
)


def normalize_ise_id(text):
    if pd.isna(text):
        return None

    text = str(text).strip()
    match = ISE_ID_PATTERN.match(text)

    if not match:
        return None

    return match.group(1)


def normalize_control_id(text):
    if pd.isna(text):
        return None

    text = str(text)
    match = CONTROL_ID_PATTERN.search(text)

    if not match:
        return None

    number = match.group(1)
    return f"CORE #{number}".upper()


def excel_col_to_index(col):
    col = col.upper().strip()
    index = 0

    for char in col:
        if not char.isalpha():
            raise ValueError(f"Invalid Excel column: {col}")
        index = index * 26 + (ord(char) - ord("A") + 1)

    return index - 1


def clean_words(text):
    if pd.isna(text):
        return []

    text = str(text).lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    words = text.split()

    words = [
        word for word in words
        if len(word) >= MIN_WORD_LENGTH and word not in STOPWORDS
    ]

    return sorted(set(words))


def compare_text(source_text, target_text):
    # Exact ISE ID handling, e.g. 10.1 vs 10.1
    source_ise_id = normalize_ise_id(source_text)
    target_ise_id = normalize_ise_id(target_text)

    if source_ise_id and target_ise_id:
        if source_ise_id == target_ise_id:
            return {
                "Match %": 1,
                "Matched Terms": source_ise_id,
                "Missing Terms": "No major missing terms",
                "Match Quality": "High",
                "Notes": "Exact ISE ID match."
            }

        return {
            "Match %": 0,
            "Matched Terms": "",
            "Missing Terms": source_ise_id,
            "Match Quality": "No Match",
            "Notes": f"ISE ID mismatch: source={source_ise_id}, target={target_ise_id}."
        }

    # Exact CORE ID handling, e.g. CORE #1.1 vs CORE 1.1
    source_control_id = normalize_control_id(source_text)
    target_control_id = normalize_control_id(target_text)

    if source_control_id and target_control_id:
        if source_control_id == target_control_id:
            return {
                "Match %": 1,
                "Matched Terms": source_control_id,
                "Missing Terms": "No major missing terms",
                "Match Quality": "High",
                "Notes": "Exact CORE control ID match."
            }

        return {
            "Match %": 0,
            "Matched Terms": "",
            "Missing Terms": source_control_id,
            "Match Quality": "No Match",
            "Notes": f"CORE control ID mismatch: source={source_control_id}, target={target_control_id}."
        }

    # Standard text comparison
    source_words = clean_words(source_text)
    target_words = set(clean_words(target_text))

    if not source_words:
        return {
            "Match %": 0,
            "Matched Terms": "",
            "Missing Terms": "No source terms",
            "Match Quality": "No Match",
            "Notes": "No usable source words to compare."
        }

    matched_words = [word for word in source_words if word in target_words]
    missing_words = [word for word in source_words if word not in target_words]
    match_percent = len(matched_words) / len(source_words)

    if match_percent >= 0.85:
        quality = "High"
        notes = "Strong alignment. Most key source terms appear in the target text."
    elif match_percent >= 0.60:
        quality = "Medium"
        notes = "Partial alignment. Review missing terms for possible gaps."
    elif match_percent > 0:
        quality = "Low"
        notes = "Weak alignment. Several key source terms are missing."
    else:
        quality = "No Match"
        notes = "No meaningful overlap found."

    return {
        "Match %": match_percent,
        "Matched Terms": ", ".join(matched_words),
        "Missing Terms": ", ".join(missing_words) if missing_words else "No major missing terms",
        "Match Quality": quality,
        "Notes": notes
    }


def safe_sheet_name(name):
    invalid_chars = r'[]:*?/\\'
    for char in invalid_chars:
        name = name.replace(char, "-")
    return name[:31]


def load_excel_column(uploaded_file, sheet_name, column_letter):
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    col_index = excel_col_to_index(column_letter)

    if col_index >= len(df.columns):
        raise IndexError(
            f"Column {column_letter} does not exist. "
            f"The selected sheet only has {len(df.columns)} columns."
        )

    return df.iloc[:, col_index]


def run_single_comparison(
    name,
    source_file,
    source_sheet,
    source_col,
    target_file,
    target_sheet,
    target_col
):
    source_series = load_excel_column(source_file, source_sheet, source_col)
    target_series = load_excel_column(target_file, target_sheet, target_col)

    max_rows = max(len(source_series), len(target_series))
    results = []

    for i in range(max_rows):
        source_text = source_series.iloc[i] if i < len(source_series) else ""
        target_text = target_series.iloc[i] if i < len(target_series) else ""

        comparison = compare_text(source_text, target_text)

        results.append({
            "Row": i + 2,
            "Comparison Name": name,
            "Source File": source_file.name,
            "Source Sheet": source_sheet,
            "Source Column": source_col,
            "Source Text": source_text,
            "Target File": target_file.name,
            "Target Sheet": target_sheet,
            "Target Column": target_col,
            "Target Text": target_text,
            **comparison
        })

    return pd.DataFrame(results)


def build_summary(all_results):
    summary_rows = []

    for comparison_name, group in all_results.groupby("Comparison Name"):
        summary_rows.append({
            "Comparison Name": comparison_name,
            "Total Rows Compared": len(group),
            "High Matches": (group["Match Quality"] == "High").sum(),
            "Medium Matches": (group["Match Quality"] == "Medium").sum(),
            "Low Matches": (group["Match Quality"] == "Low").sum(),
            "No Matches": (group["Match Quality"] == "No Match").sum(),
            "Rows Needing Review": (group["Match Quality"] != "High").sum(),
            "Average Match %": group["Match %"].mean()
        })

    return pd.DataFrame(summary_rows)


def format_workbook_openpyxl(output_buffer):
    output_buffer.seek(0)
    wb = load_workbook(output_buffer)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = wrap

        for col in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col)
            if col_letter in ["F", "J", "L", "M", "O"]:
                ws.column_dimensions[col_letter].width = 45
            elif col_letter == "K":
                ws.column_dimensions[col_letter].width = 12
            else:
                ws.column_dimensions[col_letter].width = 18

        headers = [cell.value for cell in ws[1]]
        if "Match %" in headers:
            match_col = headers.index("Match %") + 1
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=match_col).number_format = "0.00%"

    final_buffer = BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer


def create_output_workbook(results_list):
    all_results = pd.concat(results_list, ignore_index=True)
    summary = build_summary(all_results)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")

        for df in results_list:
            comparison_name = df["Comparison Name"].iloc[0]
            sheet_name = safe_sheet_name(comparison_name)
            df.to_excel(writer, index=False, sheet_name=sheet_name)

    return format_workbook_openpyxl(output), summary, all_results


def parse_sheet_value(value):
    value = str(value).strip()
    if value.isdigit():
        return int(value)
    return value


st.set_page_config(
    page_title="Excel Column Comparison Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Column Comparison Tool")

st.write(
    "Upload two Excel files, select the sheets and columns to compare, "
    "then download a clean results workbook with match %, missing terms, match quality, and notes."
)

with st.expander("How the score is calculated", expanded=False):
    st.write(
        """
        The tool compares the selected source column against the selected target column row by row.

        It handles:
        - exact ISE ID matches, such as 10.1 vs 10.1
        - exact CORE ID matches, such as CORE #1.1 vs CORE 1.1
        - text similarity using important word overlap

        For text comparison, it:
        - lowercases the text
        - removes punctuation
        - removes common stopwords
        - removes duplicate terms
        - compares source terms against target terms
        - calculates: matched source terms / total source terms
        """
    )

st.subheader("1. Upload Files")

col1, col2 = st.columns(2)

with col1:
    source_file = st.file_uploader(
        "Upload Source Excel File",
        type=["xlsx", "xls"],
        key="source_file"
    )

with col2:
    target_file = st.file_uploader(
        "Upload Target Excel File",
        type=["xlsx", "xls"],
        key="target_file"
    )

st.subheader("2. Configure Comparison")

col1, col2, col3 = st.columns(3)

with col1:
    comparison_name = st.text_input("Comparison Name", value="Comparison 1")

with col2:
    source_sheet_input = st.text_input("Source Sheet Name or 0", value="0")
    source_col = st.text_input("Source Column Letter", value="C")

with col3:
    target_sheet_input = st.text_input("Target Sheet Name or 0", value="0")
    target_col = st.text_input("Target Column Letter", value="E")

run_button = st.button("Run Comparison", type="primary")

if run_button:
    if source_file is None or target_file is None:
        st.error("Please upload both a source file and a target file.")
    else:
        try:
            source_sheet = parse_sheet_value(source_sheet_input)
            target_sheet = parse_sheet_value(target_sheet_input)

            source_file.seek(0)
            target_file.seek(0)

            result_df = run_single_comparison(
                comparison_name,
                source_file,
                source_sheet,
                source_col,
                target_file,
                target_sheet,
                target_col
            )

            output_file, summary_df, all_results_df = create_output_workbook([result_df])

            st.success("Comparison complete!")

            st.subheader("Summary")
            st.dataframe(summary_df, use_container_width=True)

            st.subheader("Preview Results")
            st.dataframe(all_results_df.head(50), use_container_width=True)

            st.download_button(
                label="Download Results Workbook",
                data=output_file,
                file_name="comparison_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
            
import html
import io
import re

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

try:
    from rapidfuzz import fuzz, process
except ImportError:  # pragma: no cover - Streamlit app fallback
    fuzz = None
    process = None


st.set_page_config(page_title="Sawtooth Translation Tool", layout="wide")
st.title("Sawtooth Translation Automation Tool")

uploaded_excel = st.file_uploader("Upload Sawtooth Export", type=["xlsx"])
uploaded_docx = st.file_uploader("Upload Translated Questionnaire", type=["docx"])


def clean_text(text, strip_html=False):
    if pd.isna(text):
        return ""

    text = html.unescape(str(text))
    text = text.replace("\xa0", " ")

    if strip_html:
        text = re.sub(r"<\s*br\s*/?>", " ", text, flags=re.I)
        text = re.sub(r"</p\s*>", " ", text, flags=re.I)
        text = re.sub(r"<[^>]+>", " ", text)

    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def strip_sawtooth_markup(text):
    text = "" if pd.isna(text) else str(text)
    text = re.sub(r"\[%.*?%\]", " ", text, flags=re.I | re.S)
    text = re.sub(r"\bVALUE\s*\([^)]*\)", " ", text, flags=re.I)
    text = re.sub(r"\bLISTLABEL\s*\([^)]*\)", " ", text, flags=re.I)
    text = re.sub(r"\bBegin\s+Unverified\s+Perl\b", " ", text, flags=re.I)
    text = re.sub(r"\bEnd\s+Unverified\b", " ", text, flags=re.I)
    return text


def strip_leading_conditions(text):
    text = clean_text(strip_sawtooth_markup(text), strip_html=True)
    while True:
        stripped = re.sub(r"^\s*\[[^\]]+\]\s*", "", text).strip()
        if stripped == text:
            return text
        text = stripped


def normalize(text):
    text = strip_leading_conditions(text).lower()
    text = re.sub(r"\[[^\]]+\]", " ", text)
    replacements = {
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "–": "-",
        "—": "-",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = text.replace("_", " ")
    text = re.sub(r"[^\w\s%./()\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def has_html_tags(text):
    return bool(re.search(r"<[A-Za-z][^>]*>|</[A-Za-z][^>]*>", str(text)))


def has_protected_markup(text):
    text = "" if pd.isna(text) else str(text)
    return bool(
        has_html_tags(text)
        or re.search(r"\[%.*?%\]", text, flags=re.I | re.S)
        or re.search(r"\b(?:VALUE|LISTLABEL)\s*\(", text, flags=re.I)
    )


def preserve_html_markup(source_text, translated_text):
    source_text = "" if pd.isna(source_text) else str(source_text)
    translated_text = clean_text(translated_text)

    if not translated_text or not has_protected_markup(source_text):
        return translated_text

    parts = re.split(r"(<[^>]+>|\[%.*?%\])", source_text, flags=re.I | re.S)
    text_indexes = [
        index
        for index, part in enumerate(parts)
        if not part.startswith("<")
        and not part.startswith("[%")
        and clean_text(strip_sawtooth_markup(part), strip_html=True)
    ]

    if not text_indexes:
        return translated_text

    first_text_index = text_indexes[0]
    first_text = parts[first_text_index]
    leading_space = re.match(r"^\s*", first_text).group(0)
    trailing_space = re.search(r"\s*$", first_text).group(0)

    parts[first_text_index] = f"{leading_space}{html.escape(translated_text, quote=False)}{trailing_space}"

    for index in text_indexes[1:]:
        text = parts[index]
        leading_space = re.match(r"^\s*", text).group(0)
        trailing_space = re.search(r"\s*$", text).group(0)
        parts[index] = f"{leading_space}{trailing_space}"

    return "".join(parts)


def looks_like_non_translation(text):
    text = clean_text(text, strip_html=True)
    if not text:
        return True

    upper = text.upper()
    if upper.startswith(
        (
            "ASK ",
            "PROGRAMMING ",
            "NOTE:",
            "QUOTA",
            "TERMINATE ",
            "CONTINUE",
            "SECTION ",
        )
    ):
        return True

    letters = re.sub(r"[^A-Za-z]+", "", text)
    return bool(letters and len(letters) > 4 and letters.isupper())


def iter_doc_blocks(doc):
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def cell_lines(cell):
    lines = []
    for paragraph in cell.paragraphs:
        for line in str(paragraph.text).splitlines():
            line = clean_text(line)
            if line:
                lines.append(line)
    return lines


def add_translation_pair(memory, source, target, origin):
    source = strip_leading_conditions(source)
    target = clean_text(target)
    key = normalize(source)

    if not key or not target:
        return

    memory.setdefault(key, {"target": target, "origin": origin, "source": source})


def build_translation_memory(doc_file):
    doc = Document(doc_file)
    memory = {}
    paragraph_run = []

    for block in iter_doc_blocks(doc):
        if isinstance(block, Paragraph):
            text = clean_text(block.text)
            if text:
                paragraph_run.append(text)
            continue

        for i in range(len(paragraph_run) - 1):
            if not looks_like_non_translation(paragraph_run[i + 1]):
                add_translation_pair(
                    memory,
                    paragraph_run[i],
                    paragraph_run[i + 1],
                    "paragraph",
                )
        paragraph_run = []

        for row in block.rows:
            for cell in row.cells:
                values = cell_lines(cell)
                for i in range(0, len(values) - 1, 2):
                    add_translation_pair(memory, values[i], values[i + 1], "table cell")

    for i in range(len(paragraph_run) - 1):
        if not looks_like_non_translation(paragraph_run[i + 1]):
            add_translation_pair(memory, paragraph_run[i], paragraph_run[i + 1], "paragraph")

    return memory


def detect_default_column(columns, exact_name=None, contains=None):
    for column in columns:
        if exact_name and str(column).strip().lower() == exact_name:
            return column
    for column in columns:
        if contains and contains in str(column).strip().lower():
            return column
    return columns[0] if columns else None


def find_translation(source_text, memory, use_fuzzy, threshold):
    key = normalize(source_text)
    if not key:
        return "", "empty", 0

    if key in memory:
        return memory[key]["target"], "exact", 100

    visible_source = clean_text(strip_sawtooth_markup(source_text), strip_html=True)
    if re.fullmatch(r"[\d.]+|[A-Z]{3}", visible_source.strip()):
        return visible_source.strip(), "copied", 100

    if use_fuzzy and process and memory:
        match = process.extractOne(key, memory.keys(), scorer=fuzz.ratio)
        if match and match[1] >= threshold:
            matched_key, score, _ = match
            return memory[matched_key]["target"], "fuzzy", int(score)

    return "", "unmatched", 0


if uploaded_excel and uploaded_docx:
    df = pd.read_excel(uploaded_excel, dtype=str, keep_default_na=False)

    if df.empty:
        st.error("The uploaded Excel file has no data rows.")
        st.stop()

    columns = list(df.columns)
    id_default = detect_default_column(columns, exact_name="id")
    source_default = detect_default_column(columns, contains="source")
    target_default = detect_default_column(columns, contains="target")

    controls = st.columns(4)
    with controls[0]:
        id_col = st.selectbox("ID column", columns, index=columns.index(id_default))
    with controls[1]:
        source_col = st.selectbox("Source column", columns, index=columns.index(source_default))
    with controls[2]:
        target_options = columns + ["Create new target column"]
        target_index = columns.index(target_default) if target_default in columns else len(columns)
        target_choice = st.selectbox("Target column", target_options, index=target_index)
    with controls[3]:
        fuzzy_threshold = st.slider("Fuzzy threshold", 90, 100, 97)

    use_fuzzy = st.checkbox("Use high-confidence fuzzy matching", value=True)
    overwrite_existing = st.checkbox("Overwrite existing target values", value=True)

    if st.button("Generate Translation File"):
        with st.spinner("Reading translated questionnaire..."):
            memory = build_translation_memory(uploaded_docx)

        target_col = "Target" if target_choice == "Create new target column" else target_choice
        if target_col not in df.columns:
            df[target_col] = ""

        source_nonblank = df[source_col].map(lambda value: clean_text(value, strip_html=True)).ne("")

        translated_values = []
        statuses = []
        scores = []

        for _, row in df.iterrows():
            existing_target = clean_text(row.get(target_col, ""))
            if existing_target and not overwrite_existing:
                translated_values.append(existing_target)
                statuses.append("kept existing")
                scores.append("")
                continue

            translation, status, score = find_translation(
                row[source_col],
                memory,
                use_fuzzy,
                fuzzy_threshold,
            )
            translated_values.append(preserve_html_markup(row[source_col], translation))
            statuses.append(status)
            scores.append(score if score else "")

        df[target_col] = translated_values

        report = pd.DataFrame(
            {
                "ID": df[id_col],
                "Source": df[source_col],
                "Target": df[target_col],
                "Status": statuses,
                "Score": scores,
            }
        )

        matched = sum(status in ("exact", "fuzzy", "copied", "kept existing") for status in statuses)
        unmatched = sum(status == "unmatched" for status in statuses)
        empty_source = int((~source_nonblank).sum())

        st.success("Translation file generated.")
        metric_cols = st.columns(5)
        metric_cols[0].metric("Excel data rows", f"{len(df):,}")
        metric_cols[1].metric("Nonblank source cells", f"{int(source_nonblank.sum()):,}")
        metric_cols[2].metric("Questionnaire pairs", f"{len(memory):,}")
        metric_cols[3].metric("Matched rows", f"{matched:,}")
        metric_cols[4].metric("Unmatched rows", f"{unmatched:,}")

        if empty_source:
            st.info(f"{empty_source:,} rows have a blank source cell; these are not counted as translatable text.")

        unmatched_preview = report[report["Status"].eq("unmatched")].head(50)
        if not unmatched_preview.empty:
            st.subheader("First unmatched rows")
            st.dataframe(unmatched_preview, use_container_width=True)

        output = io.BytesIO()
        report_output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Translated")
            report.to_excel(writer, index=False, sheet_name="Match Report")

        report.to_csv(report_output, index=False)
        output.seek(0)
        report_output.seek(0)

        st.download_button(
            "Download Translated Excel",
            data=output,
            file_name="Translated_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Download Match Report CSV",
            data=report_output,
            file_name="translation_match_report.csv",
            mime="text/csv",
        )
else:
    st.info("Upload the Sawtooth Excel export and the translated questionnaire DOCX to begin.")

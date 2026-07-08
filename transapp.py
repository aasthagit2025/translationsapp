import argparse
import io
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from docx import Document

try:
    from rapidfuzz import fuzz, process
except ImportError:
    fuzz = None
    process = None


DEFAULT_THRESHOLD = 88


@dataclass
class Candidate:
    source: str
    target: str
    method: str
    location: str


def clean_text(value):
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def strip_programming_notes(text):
    text = clean_text(text)
    text = re.sub(r"^\[[^\]]+\]\s*", "", text)
    text = re.sub(r"\s*\[(?:Exclusive|SINGLE ANSWER|MULTIPLE ANSWER|NUMERIC|RANDOMIZE)[^\]]*\]\s*$", "", text, flags=re.I)
    return clean_text(text)


def normalize(text):
    text = strip_programming_notes(text).casefold()
    text = text.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    text = re.sub(r"[_]+", " ", text)
    text = re.sub(r"[^\w\s%./'-]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def split_bilingual(text):
    text = clean_text(text)
    if " | " not in text:
        return None
    left, right = text.split(" | ", 1)
    left = strip_programming_notes(left)
    right = strip_programming_notes(right)
    if left and right:
        return left, right
    return None


def unique_cell_paragraphs(cell):
    values = []
    for paragraph in cell.paragraphs:
        for raw_line in paragraph.text.splitlines():
            value = clean_text(raw_line)
            if value:
                values.append(value)
    return values


def extract_docx_candidates(docx_file):
    doc = Document(docx_file)
    candidates = []
    paragraph_lines = []

    for idx, paragraph in enumerate(doc.paragraphs, start=1):
        text = clean_text(paragraph.text)
        if not text:
            continue
        paragraph_lines.append((idx, text))
        pair = split_bilingual(text)
        if pair:
            candidates.append(Candidate(pair[0], pair[1], "paragraph_pipe", f"paragraph {idx}"))

    for (line_no, source), (_, target) in zip(paragraph_lines, paragraph_lines[1:]):
        if source and target and normalize(source) != normalize(target):
            candidates.append(Candidate(strip_programming_notes(source), strip_programming_notes(target), "paragraph_next_line", f"paragraph {line_no}"))

    for table_no, table in enumerate(doc.tables, start=1):
        seen_cells = set()
        for row_no, row in enumerate(table.rows, start=1):
            for col_no, cell in enumerate(row.cells, start=1):
                cell_key = cell._tc
                if cell_key in seen_cells:
                    continue
                seen_cells.add(cell_key)
                location = f"table {table_no} row {row_no} col {col_no}"
                paras = unique_cell_paragraphs(cell)
                for text in paras:
                    pair = split_bilingual(text)
                    if pair:
                        candidates.append(Candidate(pair[0], pair[1], "cell_pipe", location))
                if len(paras) >= 2:
                    for source, target in zip(paras, paras[1:]):
                        candidates.append(Candidate(strip_programming_notes(source), strip_programming_notes(target), "cell_next_line", location))

    return candidates


def build_candidate_index(candidates):
    indexed = defaultdict(list)
    for candidate in candidates:
        key = normalize(candidate.source)
        if key:
            indexed[key].append(candidate)
    choices = list(indexed.keys())
    return indexed, choices


def choose_best_candidate(source, indexed, choices, threshold):
    source = clean_text(source)
    if not source:
        return "", "blank", 0, ""

    if re.fullmatch(r"[A-Z]{2,5}", source):
        return source, "copy_short_code", 100, "auto"

    key = normalize(source)
    if key in indexed:
        candidate = indexed[key][0]
        return candidate.target, candidate.method, 100, candidate.location

    simplified = normalize(strip_programming_notes(source))
    if simplified in indexed:
        candidate = indexed[simplified][0]
        return candidate.target, candidate.method, 100, candidate.location

    if choices and process is not None:
        result = process.extractOne(key, choices, scorer=fuzz.ratio)
        if result and result[1] >= threshold:
            candidate = indexed[result[0]][0]
            return candidate.target, f"fuzzy_{candidate.method}", int(result[1]), candidate.location

    if re.fullmatch(r"[\d\s.,%/_()$+-]+", source):
        return source, "copy_symbol_or_number", 100, "auto"

    return "", "unmatched", 0, ""


def detect_columns(ws):
    headers = [clean_text(ws.cell(1, col).value) for col in range(1, ws.max_column + 1)]
    id_col = None
    source_col = None
    target_cols = []
    for idx, header in enumerate(headers, start=1):
        lower = header.casefold()
        if lower == "id":
            id_col = idx
        elif "source" in lower:
            source_col = idx
        elif "target" in lower:
            target_cols.append(idx)
    if source_col is None:
        raise ValueError("Source column not found. Expected a header containing 'Source'.")
    if not target_cols:
        raise ValueError("Target column not found. Expected a header containing 'Target'.")
    return id_col, source_col, target_cols


def convert_translation(excel_file, docx_file, output_file, report_file=None, threshold=DEFAULT_THRESHOLD):
    candidates = extract_docx_candidates(docx_file)
    indexed, choices = build_candidate_index(candidates)

    wb = openpyxl.load_workbook(excel_file)
    report_rows = [["Sheet", "Row", "ID", "Source", "Target", "Method", "Score", "Location"]]
    total = matched = 0

    for ws in wb.worksheets:
        id_col, source_col, target_cols = detect_columns(ws)
        for row in range(2, ws.max_row + 1):
            source = clean_text(ws.cell(row, source_col).value)
            if not source:
                continue
            total += 1
            target, method, score, location = choose_best_candidate(source, indexed, choices, threshold)
            if target:
                matched += 1
            for target_col in target_cols:
                ws.cell(row, target_col).value = target
            row_id = clean_text(ws.cell(row, id_col).value) if id_col else ""
            report_rows.append([ws.title, row, row_id, source, target, method, score, location])

    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)

    if report_file:
        import csv
        report_file = Path(report_file)
        report_file.parent.mkdir(parents=True, exist_ok=True)
        with report_file.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            writer.writerows(report_rows)

    coverage = round((matched / total) * 100, 2) if total else 0
    return {
        "rows": total,
        "matched": matched,
        "unmatched": total - matched,
        "coverage": coverage,
        "candidates": len(candidates),
    }


def run_cli():
    parser = argparse.ArgumentParser(description="Convert a translated Word questionnaire into a Sawtooth Excel translation file.")
    parser.add_argument("excel", help="Sawtooth translation Excel file")
    parser.add_argument("docx", help="Client translated questionnaire in DOCX format")
    parser.add_argument("-o", "--output", default="translated_output.xlsx", help="Output XLSX path")
    parser.add_argument("-r", "--report", default="translation_report.csv", help="Match report CSV path")
    parser.add_argument("-t", "--threshold", type=int, default=DEFAULT_THRESHOLD, help="Fuzzy match threshold, 0-100")
    args = parser.parse_args()
    stats = convert_translation(args.excel, args.docx, args.output, args.report, args.threshold)
    print(f"Saved: {args.output}")
    print(f"Report: {args.report}")
    print(f"Matched {stats['matched']} of {stats['rows']} rows ({stats['coverage']}%).")
    print(f"Unmatched rows: {stats['unmatched']}. Candidates extracted from Word: {stats['candidates']}.")


def run_streamlit():
    import streamlit as st

    st.set_page_config(page_title="Sawtooth Translation Tool", layout="wide")
    st.title("Sawtooth Translation Automation Tool")

    uploaded_excel = st.file_uploader("Upload Sawtooth Export", type=["xlsx"])
    uploaded_docx = st.file_uploader("Upload Translated Questionnaire", type=["docx"])
    threshold = st.slider("Fuzzy match threshold", min_value=70, max_value=100, value=DEFAULT_THRESHOLD)

    if uploaded_excel and uploaded_docx and st.button("Generate Translation File"):
        with st.spinner("Extracting translations and filling workbook..."):
            xlsx_bytes = io.BytesIO(uploaded_excel.getvalue())
            docx_bytes = io.BytesIO(uploaded_docx.getvalue())
            output = io.BytesIO()
            report = io.StringIO()

            temp_output = Path("translated_output.xlsx")
            temp_report = Path("translation_report.csv")
            stats = convert_translation(xlsx_bytes, docx_bytes, temp_output, temp_report, threshold)
            output.write(temp_output.read_bytes())
            report.write(temp_report.read_text(encoding="utf-8-sig"))

        st.success(f"Matched {stats['matched']} of {stats['rows']} rows ({stats['coverage']}%). Unmatched: {stats['unmatched']}.")
        st.download_button(
            "Download translated Excel",
            data=output.getvalue(),
            file_name="translated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Download match report",
            data=report.getvalue(),
            file_name="translation_report.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        is_streamlit = get_script_run_ctx() is not None
    except Exception:
        is_streamlit = False

    if is_streamlit:
        run_streamlit()
    else:
        run_cli()

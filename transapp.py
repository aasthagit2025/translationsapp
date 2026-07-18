import html
import io
import re
from difflib import SequenceMatcher
import zipfile

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
    text = re.sub(r"\bLABEL\s*\([^)]*\)", " ", text, flags=re.I)
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
    text = text.replace("(nearsightedness)", " ")
    text = text.replace("using/ receiving", "using")
    text = text.replace("current custom-mixed atropine", "current atropine")
    text = re.sub(
        r"\bis\s+(?:sgd|twd|thb|krw|myr|php)(?:\s*/\s*(?:sgd|twd|thb|krw|myr|php))*\s*",
        "is ",
        text,
    )
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
    text = re.sub(r"\byour child s\b", "your", text)
    text = re.sub(r"[^\w\s%./()\-]+", " ", text)
    text = re.sub(r"^\s*[a-z]\.\s+", " ", text)
    text = re.sub(r"\s+-\s+", " ", text)
    text = re.sub(r"\s+", " ", text)
    endpoint = re.match(
        r"^((?:very|extremely|not|no)\b.+?)\s+(\d{1,2})$",
        text,
    )
    if endpoint:
        text = f"{endpoint.group(2)} {endpoint.group(1)}"
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

    if len(text_indexes) > 1 and not re.search(r"\[%.*?%\]", source_text, flags=re.I | re.S):
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


def translate_plain_text(source_text, memory, use_fuzzy, threshold):
    key = normalize(source_text)
    if not key:
        if has_protected_markup(source_text):
            return str(source_text), "copied_markup", 100
        return "", "empty", 0

    if key in memory:
        if memory[key].get("ambiguous"):
            return "", "ambiguous", 0
        return memory[key]["target"], "exact", 100

    visible_source = clean_text(strip_sawtooth_markup(source_text), strip_html=True)
    if re.fullmatch(r"[\d.]+|[A-Z]{3}", visible_source.strip()):
        return visible_source.strip(), "copied", 100

    if re.search(r"<script\b", str(source_text), flags=re.I):
        return str(source_text), "copied_markup", 100

    if use_fuzzy and process and memory:
        match = process.extractOne(key, memory.keys(), scorer=fuzz.ratio)
        if match and match[1] >= threshold:
            matched_key, score, _ = match
            if memory[matched_key].get("ambiguous"):
                return "", "ambiguous", 0
            return memory[matched_key]["target"], "fuzzy", int(score)
        if (
            match
            and len(key) > 80
            and match[1] >= 60
            and key[:35] == match[0][:35]
        ):
            if memory[match[0]].get("ambiguous"):
                return "", "ambiguous", 0
            return memory[match[0]]["target"], "fuzzy_long", int(match[1])

    return "", "unmatched", 0


def translate_from_contained_phrases(source_text, memory):
    key = normalize(source_text)
    if len(key) < 40:
        return "", "unmatched", 0

    candidates = []
    for memory_key, item in memory.items():
        if len(memory_key) < 40:
            continue
        position = key.find(memory_key)
        if position >= 0:
            candidates.append((position, position + len(memory_key), len(memory_key), item["target"]))

    if not candidates:
        return "", "unmatched", 0

    candidates.sort(key=lambda item: (item[0], -item[2]))
    selected = []
    occupied_until = -1
    for start, end, _, target in candidates:
        if start >= occupied_until:
            selected.append(target)
            occupied_until = end

    if selected:
        return " ".join(selected), "composite", 100

    return "", "unmatched", 0


def is_meaningful_untranslated_text(text):
    text = clean_text(strip_sawtooth_markup(text), strip_html=True)
    text = re.sub(r"https?://\S+", " ", text, flags=re.I)
    letters = re.sub(r"[^A-Za-z]+", "", text)
    return len(letters) > 4


def translate_composite_markup(source_text, memory, use_fuzzy, threshold):
    source_text = "" if pd.isna(source_text) else str(source_text)
    if not has_protected_markup(source_text):
        return "", "unmatched", 0

    parts = re.split(r"(<br\s*/?>|\[%.*?%\])", source_text, flags=re.I | re.S)
    translated_count = 0
    untranslated_count = 0
    scores = []

    for index, part in enumerate(parts):
        if not part or re.fullmatch(r"<br\s*/?>", part, flags=re.I) or part.startswith("[%"):
            continue

        leading_space = re.match(r"^\s*", part).group(0)
        trailing_space = re.search(r"\s*$", part).group(0)
        core = part[len(leading_space) : len(part) - len(trailing_space)]
        translation, status, score = translate_plain_text(
            core,
            memory,
            use_fuzzy,
            threshold,
        )
        if not translation:
            translation, status, score = translate_from_contained_phrases(core, memory)
        if translation:
            translated_part = preserve_html_markup(core, translation)
            parts[index] = f"{leading_space}{translated_part}{trailing_space}"
            translated_count += 1
            if score:
                scores.append(score)
        elif is_meaningful_untranslated_text(core):
            untranslated_count += 1

    if translated_count and untranslated_count <= translated_count:
        score = min(scores) if scores else 100
        return "".join(parts), "composite", score

    return "", "unmatched", 0


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

    def comparable_target(value):
        value = normalize(value)
        value = re.sub(r"(?<=[\u0600-\u06FF])\s+(?=[\u0600-\u06FF])", "", value)
        return value

    existing = memory.get(key)
    if existing:
        if comparable_target(existing["target"]) != comparable_target(target):
            existing["ambiguous"] = True
            existing["alternatives"] = sorted(
                set(existing.get("alternatives", [existing["target"]]) + [target])
            )
        elif len(target) < len(existing["target"]):
            existing["target"] = target
        return

    memory[key] = {
        "target": target,
        "origin": origin,
        "source": source,
        "ambiguous": False,
        "alternatives": [target],
    }


def contains_urdu(text):
    return bool(re.search(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]", str(text)))


def add_mixed_script_pairs_from_text(memory, text, origin, allowed_source_map=None):
    text = clean_text(text, strip_html=True)
    if not text or not contains_urdu(text):
        return

    urdu_block_pattern = re.compile(
        r"(?:[\d\s,.-]+)?"
        r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]"
        r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\s،؛؟ءآأؤإئابپتٹثجچحخدڈذرڑزژسشصضطظعغفقکگلمنںوہھہیےَُِّْٰ]*"
    )

    for match in urdu_block_pattern.finditer(text):
        target = clean_text(match.group(0))
        if not target:
            continue

        prefix = text[max(0, match.start() - 220) : match.start()]
        source_tokens = re.findall(
            r"[A-Za-z][A-Za-z0-9’'&/().,\-]*",
            prefix,
        )
        if not source_tokens:
            continue

        stopwords = {
            "code",
            "action",
            "continue",
            "close",
            "terminate",
            "region",
            "city",
            "area",
            "instruction",
            "instructions",
            "scripter",
        }
        source_tokens = [token for token in source_tokens if token.lower() not in stopwords]

        candidates = []
        prefix_key = normalize(prefix)
        if allowed_source_map:
            for allowed_key, allowed_source in allowed_source_map.items():
                if len(allowed_key) < 2:
                    continue
                if prefix_key.endswith(allowed_key):
                    candidates.append(allowed_source)

        for length in range(1, min(14, len(source_tokens)) + 1):
            source = " ".join(source_tokens[-length:])
            if len(source) < 2:
                continue
            if allowed_source_map is not None and normalize(source) not in allowed_source_map:
                continue
            candidates.append(source)

        if candidates:
            source = max(candidates, key=lambda value: len(normalize(value)))
            add_translation_pair(memory, source, target, f"{origin} mixed script")


def add_mixed_script_pairs_from_docx_xml(memory, doc_file, allowed_source_map=None):
    try:
        with zipfile.ZipFile(doc_file) as package:
            xml_text = package.read("word/document.xml").decode("utf-8", "ignore")
    except Exception:
        return

    xml_text = re.sub(r"<[^>]+>", " ", xml_text)
    xml_text = html.unescape(xml_text)
    xml_text = clean_text(xml_text)
    add_mixed_script_pairs_from_text(memory, xml_text, "docx xml", allowed_source_map)
    add_allowed_source_pairs_from_text(memory, xml_text, allowed_source_map, "docx xml exact source")


def source_to_flexible_regex(source):
    escaped = re.escape(clean_text(source, strip_html=True))
    escaped = re.sub(r"\\\s+", r"\\s+", escaped)
    escaped = escaped.replace("\\–", "[–—-]").replace("\\-", "[–—-]")
    escaped = escaped.replace(",", r"\s*,\s*")
    return escaped


def extract_following_urdu_target(tail):
    tail = tail[:260]
    code_break = re.search(r"\s+\d{1,3}\s+(?=[A-Za-z])", tail)
    if code_break:
        tail = tail[: code_break.start()]

    match = re.match(
        r"\s*((?:[\d\s,.-]+)?[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]"
        r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\d\s,.-،؛؟ءآأؤإئابپتٹثجچحخدڈذرڑزژسشصضطظعغفقکگلمنںوہھہیےَُِّْٰ]*)",
        tail,
    )
    if not match:
        return ""

    target = clean_text(match.group(1))
    return target if contains_urdu(target) else ""


def add_allowed_source_pairs_from_text(memory, text, allowed_source_map, origin):
    if not allowed_source_map or not contains_urdu(text):
        return

    source_values = sorted(
        set(allowed_source_map.values()),
        key=lambda value: len(clean_text(value, strip_html=True)),
        reverse=True,
    )
    for source in source_values:
        clean_source = clean_text(source, strip_html=True)
        if len(clean_source) < 2:
            continue
        pattern = source_to_flexible_regex(clean_source)
        for match in re.finditer(pattern, text, flags=re.I):
            target = extract_following_urdu_target(text[match.end() :])
            if target:
                add_translation_pair(memory, clean_source, target, origin)
                break


def add_common_urdu_pairs(memory):
    if not any(contains_urdu(item["target"]) for item in memory.values()):
        return

    common_pairs = {
        "English": "انگریزی",
        "Urdu": "اردو",
        "Central Punjab": "وسطی پنجاب",
        "North": "شمال",
        "Sindh & Baluchistan": "سندھ اور بلوچستان",
        "South Punjab": "جنوبی پنجاب",
    }
    for source, target in common_pairs.items():
        add_translation_pair(memory, source, target, "common Urdu fallback")


def add_derived_scale_pairs(memory):
    additions = {}
    endpoint_additions = {}
    for key, item in memory.items():
        source_match = re.match(r"^(\d{1,2})\s+(.+)$", key)
        if source_match:
            number, source_label = source_match.groups()
            target_match = re.match(
                rf"^\s*{re.escape(number)}\s*[–\-]?\s*(.+?)\s*$",
                item["target"],
            )
            if target_match:
                additions.setdefault(
                    normalize(source_label),
                    {
                        "target": target_match.group(1),
                        "origin": f"{item['origin']} scale label",
                        "source": source_label,
                    },
                )

        endpoints = re.search(
            r"where\s+1\s*=?\s*(.+?)\s+and\s+7\s*=?\s*(.+?)(?:,|\.|\?|$)",
            item["source"],
            flags=re.I,
        )
        target_endpoints = re.search(
            r"kung\s+saan.*?1\s*(?:=|ay nangangahulugang)\s*\"?(.+?)\"?\s+at\s+(?:ang\s+)?7\s*(?:=|ay nangangahulugang)\s*\"?(.+?)\"?(?:,|\.|$)",
            item["target"],
            flags=re.I,
        )
        if endpoints and target_endpoints:
            endpoint_additions.setdefault(
                normalize(endpoints.group(1).lstrip("= ")),
                {
                    "target": target_endpoints.group(1).strip(),
                    "origin": f"{item['origin']} scale endpoint",
                    "source": endpoints.group(1),
                },
            )
            endpoint_additions.setdefault(
                normalize(endpoints.group(2).lstrip("= ")),
                {
                    "target": target_endpoints.group(2).strip(),
                    "origin": f"{item['origin']} scale endpoint",
                    "source": endpoints.group(2),
                },
            )

    for key, value in endpoint_additions.items():
        additions.setdefault(key, value)
    memory.update(additions)


def build_translation_memory(doc_file, source_values=None):
    doc = Document(doc_file)
    memory = {}
    paragraph_run = []
    allowed_source_map = None
    if source_values is not None:
        allowed_source_map = {
            normalize(value): clean_text(value, strip_html=True)
            for value in source_values
            if clean_text(value, strip_html=True)
        }

    for block in iter_doc_blocks(doc):
        if isinstance(block, Paragraph):
            text = clean_text(block.text)
            if text:
                paragraph_run.append(text)
                add_mixed_script_pairs_from_text(memory, text, "paragraph", allowed_source_map)
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
                add_mixed_script_pairs_from_text(memory, cell.text, "table cell", allowed_source_map)
                for i in range(0, len(values) - 1, 2):
                    add_translation_pair(memory, values[i], values[i + 1], "table cell")

    for i in range(len(paragraph_run) - 1):
        if not looks_like_non_translation(paragraph_run[i + 1]):
            add_translation_pair(memory, paragraph_run[i], paragraph_run[i + 1], "paragraph")

    add_derived_scale_pairs(memory)
    add_mixed_script_pairs_from_docx_xml(memory, doc_file, allowed_source_map)
    add_common_urdu_pairs(memory)
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
    if has_protected_markup(source_text) and re.search(r"<br\s*/?>", str(source_text), flags=re.I):
        translation, status, score = translate_composite_markup(
            source_text,
            memory,
            use_fuzzy,
            threshold,
        )
        if translation:
            return translation, status, score

    translation, status, score = translate_plain_text(
        source_text,
        memory,
        use_fuzzy,
        threshold,
    )
    if translation:
        return translation, status, score

    if status == "empty":
        return translation, status, score

    return translate_composite_markup(source_text, memory, use_fuzzy, threshold)


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
        fuzzy_threshold = st.slider("Fuzzy threshold", 90, 100, 95)

    use_fuzzy = st.checkbox("Use high-confidence fuzzy matching", value=False)
    overwrite_existing = st.checkbox("Overwrite existing target values", value=True)

    if st.button("Generate Translation File"):
        with st.spinner("Reading translated questionnaire..."):
            memory = build_translation_memory(uploaded_docx, df[source_col])

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
            if status in ("composite", "copied_markup"):
                translated_values.append(translation)
            else:
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

        matched_statuses = ("exact", "fuzzy", "fuzzy_long", "copied", "copied_markup", "composite", "kept existing")
        matched = sum(status in matched_statuses for status in statuses)
        unmatched = sum(status not in matched_statuses and status != "empty" for status in statuses)
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

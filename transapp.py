import streamlit as st
import pandas as pd
import re
import io

from docx import Document
from rapidfuzz import fuzz

# =====================================================
# CONFIG
# =====================================================

FUZZY_THRESHOLD = 92

# =====================================================
# HELPERS
# =====================================================

def normalize_text(text):

    if pd.isna(text):
        return ""

    text = str(text)

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def clean_for_match(text):

    text = normalize_text(text)

    text = text.lower()

    text = re.sub(r"[^\w\s]", "", text)

    return text.strip()


def contains_latin(text):

    return bool(re.search(r"[A-Za-z]", str(text)))


# =====================================================
# DOCX READER
# =====================================================

def read_docx_content(uploaded_docx):

    doc = Document(uploaded_docx)

    content = []

    # Paragraphs
    for p in doc.paragraphs:

        txt = normalize_text(p.text)

        if txt:
            content.append(txt)

    # Tables
    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                txt = normalize_text(cell.text)

                if txt:
                    content.append(txt)

    return content


# =====================================================
# TRANSLATION DICTIONARY
# =====================================================

def build_translation_dictionary(content):

    translations = {}

    for i in range(len(content) - 1):

        current = normalize_text(content[i])
        nxt = normalize_text(content[i + 1])

        if not current:
            continue

        if not nxt:
            continue

        if current == nxt:
            continue

        if len(current) < 2:
            continue

        if len(nxt) < 2:
            continue

        if current.isnumeric():
            continue

        if nxt.isnumeric():
            continue

        if contains_latin(current):

            key = clean_for_match(current)

            if key not in translations:

                translations[key] = nxt

    return translations


# =====================================================
# FUZZY MATCH
# =====================================================

def fuzzy_lookup(source_key, translation_dict):

    best_score = 0
    best_translation = ""

    for k, v in translation_dict.items():

        score = fuzz.ratio(source_key, k)

        if score > best_score:

            best_score = score
            best_translation = v

    if best_score >= FUZZY_THRESHOLD:

        return best_translation

    return ""


# =====================================================
# STREAMLIT UI
# =====================================================

st.set_page_config(
    page_title="Survey Translation Tool",
    layout="wide"
)

st.title("Survey Translation Automation Tool")

st.markdown(
    "Upload Sawtooth Export and translated Word questionnaire."
)

uploaded_excel = st.file_uploader(
    "Upload Sawtooth Export (.xlsx)",
    type=["xlsx"]
)

uploaded_docx = st.file_uploader(
    "Upload Translated Questionnaire (.docx)",
    type=["docx"]
)

# =====================================================
# PROCESS
# =====================================================

if uploaded_excel and uploaded_docx:

    if st.button("Generate Translation File"):

        with st.spinner("Reading questionnaire..."):

            content = read_docx_content(uploaded_docx)

            translation_dict = build_translation_dictionary(
                content
            )

        st.success(
            f"{len(translation_dict)} translation pairs extracted"
        )

        # ==========================================
        # READ EXCEL
        # ==========================================

        df = pd.read_excel(uploaded_excel)

        source_col = None
        target_col = None
        id_col = None

        for col in df.columns:

            lc = col.lower()

            if lc == "id":
                id_col = col

            if "source" in lc:
                source_col = col

            if "target" in lc:
                target_col = col

        if source_col is None:

            st.error("Source column not found")

            st.stop()

        if target_col is None:

            st.error("Target column not found")

            st.stop()

        # ==========================================
        # TRANSLATE
        # ==========================================

        translated = []
        match_types = []
        unmatched = []

        translated_count = 0

        for idx, row in df.iterrows():

            source_text = normalize_text(
                row[source_col]
            )

            existing_target = normalize_text(
                row[target_col]
            )

            if existing_target:

                translated.append(existing_target)

                match_types.append(
                    "Already Exists"
                )

                translated_count += 1

                continue

            source_key = clean_for_match(
                source_text
            )

            result = ""
            match_type = ""

            # ----------------------
            # Exact Match
            # ----------------------

            if source_key in translation_dict:

                result = translation_dict[source_key]

                match_type = "Exact"

            # ----------------------
            # Fuzzy Match
            # ----------------------

            if result == "":

                fuzzy_result = fuzzy_lookup(
                    source_key,
                    translation_dict
                )

                if fuzzy_result:

                    result = fuzzy_result

                    match_type = "Fuzzy"

            # ----------------------
            # Unmatched
            # ----------------------

            if result:

                translated_count += 1

            else:

                unmatched.append({

                    "Row": idx + 2,

                    "ID":
                    row[id_col]
                    if id_col else "",

                    "Source":
                    source_text

                })

                match_type = "Unmatched"

            translated.append(result)

            match_types.append(match_type)

        # ==========================================
        # OUTPUT
        # ==========================================

        df[target_col] = translated

        df["Match Type"] = match_types

        unmatched_df = pd.DataFrame(
            unmatched
        )

        coverage = round(
            translated_count
            / len(df)
            * 100,
            2
        )

        summary_df = pd.DataFrame({

            "Metric": [

                "Total Rows",

                "Translated Rows",

                "Manual Review",

                "Coverage %"

            ],

            "Value": [

                len(df),

                translated_count,

                len(unmatched),

                coverage

            ]
        })

        # ==========================================
        # DOWNLOAD FILE
        # ==========================================

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
        ) as writer:

            df.to_excel(
                writer,
                sheet_name="Translations",
                index=False
            )

            unmatched_df.to_excel(
                writer,
                sheet_name="Unmatched",
                index=False
            )

            summary_df.to_excel(
                writer,
                sheet_name="Summary",
                index=False
            )

        output.seek(0)

        st.subheader("Coverage Report")

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Total Rows",
            len(df)
        )

        c2.metric(
            "Translated",
            translated_count
        )

        c3.metric(
            "Manual Review",
            len(unmatched)
        )

        c4.metric(
            "Coverage %",
            coverage
        )

        st.download_button(
            label="Download Output Workbook",
            data=output,
            file_name="Translated_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
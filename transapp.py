import streamlit as st
import pandas as pd
import io
import re
from docx import Document

# ==========================================
# CONFIG
# ==========================================

st.set_page_config(
    page_title="Translation Automation Tool",
    layout="wide"
)

st.title("Survey Translation Automation Tool")

# ==========================================
# INPUTS
# ==========================================

uploaded_excel = st.file_uploader(
    "Upload Sawtooth Export",
    type=["xlsx"]
)

uploaded_docx = st.file_uploader(
    "Upload Questionnaire",
    type=["docx"]
)

language = st.selectbox(
    "Translation Language",
    [
        "Hindi",
        "Marathi",
        "Tamil",
        "Telugu",
        "Kannada",
        "Malayalam",
        "Gujarati",
        "Punjabi",
        "Bengali",
        "Odia",
        "Filipino",
        "Japanese",
        "Korean",
        "Arabic",
        "French",
        "German",
        "Spanish"
    ]
)

translation_color = st.selectbox(
    "Translation Color",
    [
        "Green",
        "Blue",
        "Red",
        "Purple",
        "Black"
    ]
)

# ==========================================
# HELPERS
# ==========================================

def normalize(txt):

    if pd.isna(txt):
        return ""

    txt = str(txt)

    txt = txt.replace("\n", " ")

    txt = txt.replace("\r", " ")

    txt = re.sub(r"\s+", " ", txt)

    return txt.strip()


def clean(txt):

    txt = normalize(txt)

    txt = txt.lower()

    txt = re.sub(r"[^\w\s]", "", txt)

    return txt.strip()


# ==========================================
# READ WORD
# ==========================================

def read_word(doc_file):

    doc = Document(doc_file)

    content = []

    # paragraphs

    for p in doc.paragraphs:

        txt = normalize(p.text)

        if txt:

            content.append(txt)

    # tables

    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                txt = normalize(cell.text)

                if txt:

                    content.append(txt)

    return content


# ==========================================
# BUILD DICTIONARY
# ==========================================

def build_dictionary(content):

    translations = {}

    for i in range(len(content)-1):

        en = normalize(content[i])

        tr = normalize(content[i+1])

        if not en:
            continue

        if not tr:
            continue

        if en == tr:
            continue

        if len(en) < 2:
            continue

        if len(tr) < 2:
            continue

        key = clean(en)

        if key not in translations:

            translations[key] = tr

    return translations


# ==========================================
# PROCESS
# ==========================================

if uploaded_excel and uploaded_docx:

    if st.button("Generate Translation File"):

        with st.spinner("Reading questionnaire..."):

            word_content = read_word(
                uploaded_docx
            )

            translation_dict = build_dictionary(
                word_content
            )

        df = pd.read_excel(
            uploaded_excel
        )

        source_col = None
        target_col = None

        for col in df.columns:

            lc = col.lower()

            if "source" in lc:

                source_col = col

            if "target" in lc:

                target_col = col

        if source_col is None:

            st.error(
                "Source column not found"
            )

            st.stop()

        if target_col is None:

            st.error(
                "Target column not found"
            )

            st.stop()

        translated = []

        for _, row in df.iterrows():

            source_text = normalize(
                row[source_col]
            )

            key = clean(
                source_text
            )

            target = ""

            if key in translation_dict:

                target = translation_dict[key]

            translated.append(
                target
            )

        df[target_col] = translated

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
        ) as writer:

            df.to_excel(
                writer,
                index=False
            )

        output.seek(0)

        st.success(
            "Translation file generated"
        )

        st.download_button(
            "Download Output",
            output,
            file_name="Translated_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
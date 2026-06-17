import streamlit as st
import pandas as pd
import io
import re
from docx import Document

# ==================================================
# UI
# ==================================================

st.set_page_config(
    page_title="Sawtooth Translation Tool",
    layout="wide"
)

st.title("Sawtooth Translation Automation Tool")

uploaded_excel = st.file_uploader(
    "Upload Sawtooth Export",
    type=["xlsx"]
)

uploaded_docx = st.file_uploader(
    "Upload Questionnaire",
    type=["docx"]
)

# ==================================================
# HELPERS
# ==================================================

def clean_text(text):

    if pd.isna(text):
        return ""

    text = str(text)

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def normalize(text):

    text = clean_text(text)

    text = text.lower()

    text = re.sub(r"[^\w\s]", "", text)

    return text.strip()


# ==================================================
# READ WORD
# ==================================================

def read_docx(doc_file):

    doc = Document(doc_file)

    content = []

    for para in doc.paragraphs:

        txt = clean_text(para.text)

        if txt:

            content.append(txt)

    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                txt = clean_text(cell.text)

                if txt:

                    content.append(txt)

    return content


# ==================================================
# BUILD QUESTION BLOCKS
# ==================================================

def build_blocks(content):

    blocks = {}

    current_block = None

    pattern = re.compile(
        r"^([A-Z]+\d+)\."
    )

    for line in content:

        m = pattern.match(line)

        if m:

            current_block = m.group(1)

            blocks[current_block] = []

        if current_block:

            blocks[current_block].append(line)

    return blocks


# ==================================================
# GET QUESTION ID
# ==================================================

def get_question_id(id_value):

    if pd.isna(id_value):

        return ""

    id_value = str(id_value)

    m = re.match(
        r"([A-Z]+\d+)",
        id_value,
        re.I
    )

    if m:

        return m.group(1).upper()

    return ""


# ==================================================
# FIND TRANSLATION
# ==================================================

def find_translation(
        source_text,
        block_lines
):

    source_norm = normalize(
        source_text
    )

    for i in range(
        len(block_lines)-1
    ):

        current = clean_text(
            block_lines[i]
        )

        nxt = clean_text(
            block_lines[i+1]
        )

        if normalize(current) == source_norm:

            return nxt

    return ""


# ==================================================
# PROCESS
# ==================================================

if uploaded_excel and uploaded_docx:

    if st.button(
        "Generate Translation File"
    ):

        with st.spinner(
            "Reading Questionnaire..."
        ):

            content = read_docx(
                uploaded_docx
            )

            blocks = build_blocks(
                content
            )

        st.success(
            f"{len(blocks)} question blocks found"
        )

        df = pd.read_excel(
            uploaded_excel
        )

        # detect columns

        id_col = None
        source_col = None
        target_col = None

        for col in df.columns:

            lc = col.lower()

            if lc == "id":

                id_col = col

            elif "source" in lc:

                source_col = col

            elif "target" in lc:

                target_col = col

        if not id_col:

            st.error(
                "ID column not found"
            )
            st.stop()

        if not source_col:

            st.error(
                "Source column not found"
            )
            st.stop()

        if not target_col:

            st.error(
                "Target column not found"
            )
            st.stop()

        translated = []

        unmatched = 0

        for _, row in df.iterrows():

            qid = get_question_id(
                row[id_col]
            )

            source_text = clean_text(
                row[source_col]
            )

            target = ""

            if (
                qid in blocks
                and source_text
            ):

                target = find_translation(
                    source_text,
                    blocks[qid]
                )

            if not target:

                unmatched += 1

            translated.append(
                target
            )

        df[target_col] = translated

        coverage = round(
            (
                (len(df)-unmatched)
                /
                len(df)
            ) * 100,
            2
        )

        st.write(
            f"Coverage: {coverage}%"
        )

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

        st.download_button(
            "Download Translated File",
            data=output,
            file_name="Translated_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
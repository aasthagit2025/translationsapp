import pandas as pd
import re
from docx import Document
from rapidfuzz import fuzz
from collections import defaultdict

##########################################################
# CONFIG
##########################################################

SAWTOOTH_FILE = r"Input/Sawtooth_Export.xlsx"
WORD_FILE = r"Input/QRE_Translated.docx"

OUTPUT_FILE = r"Output/Translated_Output.xlsx"
UNMATCHED_FILE = r"Output/Unmatched.xlsx"
REPORT_FILE = r"Output/Translation_Report.xlsx"

FUZZY_THRESHOLD = 95

##########################################################
# TEXT CLEANER
##########################################################

def normalize_text(text):

    if pd.isna(text):
        return ""

    text = str(text)

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")

    text = re.sub(r"\s+", " ", text)

    text = text.strip()

    return text

def clean_for_match(text):

    text = normalize_text(text)

    text = text.lower()

    text = re.sub(r"[^\w\s]", "", text)

    return text

##########################################################
# QUESTION ID EXTRACTOR
##########################################################

def extract_qid(id_text):

    if pd.isna(id_text):
        return None

    id_text = str(id_text)

    patterns = [

        r"'([A-Z]+\d+[a-zA-Z]*)List'",
        r"'([A-Z]+\d+[a-zA-Z]*)Grid'",
        r"'([A-Z]+\d+[a-zA-Z]*)Question'",
        r"'([A-Z]+\d+[a-zA-Z]*)'",
    ]

    for p in patterns:

        m = re.search(p, id_text)

        if m:
            return m.group(1)

    return None

##########################################################
# READ DOCX
##########################################################

def read_docx_content(doc_path):

    doc = Document(doc_path)

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

##########################################################
# BUILD TRANSLATION PAIRS
##########################################################

def build_translation_dictionary(content):

    translations = {}

    for i in range(len(content)-1):

        en = content[i]
        tr = content[i+1]

        if en == tr:
            continue

        if len(en) < 2:
            continue

        if len(tr) < 2:
            continue

        key = clean_for_match(en)

        if key not in translations:

            translations[key] = tr

    return translations

##########################################################
# FUZZY LOOKUP
##########################################################

def fuzzy_lookup(source, translation_dict):

    best_score = 0
    best_translation = ""

    for k, v in translation_dict.items():

        score = fuzz.ratio(source, k)

        if score > best_score:

            best_score = score
            best_translation = v

    if best_score >= FUZZY_THRESHOLD:

        return best_translation

    return ""

##########################################################
# LOAD WORD QRE
##########################################################

print("Reading Word file...")

content = read_docx_content(WORD_FILE)

translation_dict = build_translation_dictionary(content)

print(
    f"Translation pairs extracted: "
    f"{len(translation_dict)}"
)

##########################################################
# LOAD SAWTOOTH EXPORT
##########################################################

print("Reading Sawtooth export...")

df = pd.read_excel(SAWTOOTH_FILE)

##########################################################
# DETECT COLUMNS
##########################################################

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

if not source_col:
    raise Exception("Source column not found")

if not target_col:
    raise Exception("Target column not found")

##########################################################
# TRANSLATION
##########################################################

translated_count = 0

unmatched_rows = []

translations = []

for idx, row in df.iterrows():

    source_text = normalize_text(
        row[source_col]
    )

    existing_target = normalize_text(
        row[target_col]
    )

    if existing_target:

        translations.append(existing_target)
        continue

    source_key = clean_for_match(source_text)

    translated = ""

    ######################################################
    # EXACT MATCH
    ######################################################

    if source_key in translation_dict:

        translated = translation_dict[source_key]

    ######################################################
    # FUZZY MATCH
    ######################################################

    if translated == "":

        translated = fuzzy_lookup(
            source_key,
            translation_dict
        )

    ######################################################
    # SAVE RESULT
    ######################################################

    if translated:

        translated_count += 1

    else:

        unmatched_rows.append({

            "Row": idx + 2,

            "ID": row[id_col]
            if id_col else "",

            "Source": source_text
        })

    translations.append(translated)

##########################################################
# WRITE OUTPUT
##########################################################

df[target_col] = translations

df.to_excel(
    OUTPUT_FILE,
    index=False
)

##########################################################
# UNMATCHED REPORT
##########################################################

unmatched_df = pd.DataFrame(
    unmatched_rows
)

unmatched_df.to_excel(
    UNMATCHED_FILE,
    index=False
)

##########################################################
# SUMMARY REPORT
##########################################################

coverage = round(
    translated_count /
    len(df) * 100,
    2
)

report_df = pd.DataFrame({

    "Metric":[
        "Total Rows",
        "Translated Rows",
        "Manual Review",
        "Coverage %"
    ],

    "Value":[
        len(df),
        translated_count,
        len(unmatched_rows),
        coverage
    ]
})

report_df.to_excel(
    REPORT_FILE,
    index=False
)

##########################################################
# DONE
##########################################################

print("="*50)
print("Translation Complete")
print("="*50)

print(f"Total Rows     : {len(df)}")
print(f"Translated     : {translated_count}")
print(f"Manual Review  : {len(unmatched_rows)}")
print(f"Coverage       : {coverage}%")

print()
print("Files Generated:")
print(OUTPUT_FILE)
print(UNMATCHED_FILE)
print(REPORT_FILE)
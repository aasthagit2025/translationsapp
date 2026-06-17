import streamlit as st
import pandas as pd
from docx import Document
from rapidfuzz import fuzz
import io
import re

st.set_page_config(page_title="Final Sawtooth Layout Automator", layout="centered")

st.title("🎯 Final Sawtooth Translation Automator")
st.write("Strictly correlates Question ID -> Source Text Token -> Clean Translation Mapping.")

# --- Core Matrix Alignment Logic ---
def parse_word_by_question_blocks(docx_file):
    """
    Chronologically splits the questionnaire docx file into strict dictionary sections.
    Key: Question Code (e.g., INTRO, S1, S2, Q1)
    Value: List of strings (options, text blocks) belonging exclusively to that question.
    """
    doc = Document(docx_file)
    sections = {"GLOBAL": []}
    current_section = "GLOBAL"
    
    all_lines = []
    # Extract paragraphs
    for para in doc.paragraphs:
        txt = para.text.strip()
        if txt:
            all_lines.append(txt)
    # Extract tables
    for table in doc.tables:
        for row in table.rows:
            cells_txt = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells_txt:
                all_lines.append(" | ".join(cells_txt))
                
    # Strict regex to identify true question headers (e.g., Intro., S1., S2., Q5., [S1])
    # Looks for S, Q, D or Intro followed by digits/letters at the beginning of a line
    q_pattern = re.compile(r'^(?:\[)?(INTRO|[SQD]\d+[a-zA-Z0-9]*)(?:\])?[\.\s\:]', re.IGNORECASE)
    
    for line in all_lines:
        match = q_pattern.match(line)
        if match:
            current_section = match.group(1).upper()
            if current_section not in sections:
                sections[current_section] = []
        
        sections[current_section].append(line)
        sections["GLOBAL"].append(line)
        
    return sections

def extract_strict_translation(source_text, section_lines):
    """
    Looks inside the isolated question block lines for a structured line (comma or pipe separated)
    where one part exactly matches the English source_text, then returns the corresponding translation.
    """
    src_clean = source_text.strip().lower()
    
    for line in section_lines:
        # Split line by common survey layout delimiters
        parts = []
        if '|' in line:
            parts = [p.strip() for p in line.split('|')]
        elif ',' in line:
            parts = [p.strip() for p in line.split(',')]
        else:
            # Check sequential lines layout if not split on a single line
            continue
            
        if len(parts) >= 2:
            # Step 2: Identify where the English source text matches
            for idx, part in enumerate(parts):
                part_clean = part.strip().lower()
                
                # Check for exact or highly precise token match
                if part_clean == src_clean or fuzz.ratio(part_clean, src_clean) > 96:
                    # Find the translated string next to it or across the tokens
                    # Usually: [English, Translation, Code] or [English, Translation]
                    for target_idx, candidate in enumerate(parts):
                        if target_idx != idx and not candidate.isdigit():
                            cand_clean = candidate.strip()
                            # Ensure it isn't identical to the source English text
                            if cand_clean.lower() != src_clean:
                                return cand_clean
                                
    # Fallback to consecutive line pairing check within the section
    for i in range(len(section_lines) - 1):
        if section_lines[i].strip().lower() == src_clean:
            next_line = section_lines[i+1].strip()
            # Verify the next line isn't a routing command or structural label
            if next_line and not next_line.startswith('[') and not next_line.isdigit():
                return next_line
                
    return None

def identify_question_code(sawtooth_id):
    """
    Extracts the explicit question identifier label from the Sawtooth ID string.
    e.g., "List 'IntroList' - Item 1 - Display Text" -> "INTRO"
    e.g., "Question 'S1' - Body Text" -> "S1"
    """
    sawtooth_id = str(sawtooth_id)
    # Target strings inside single quotes first
    match = re.search(r"'([a-zA-Z0-9]+)", sawtooth_id)
    if match:
        code = match.group(1)
        # Strip trailing functional suffixes like 'List' or 'Term'
        code = re.sub(r'(List|Term|Qnr|Labels)$', '', code, flags=re.IGNORECASE)
        return code.upper()
    
    # Fallback if no single quotes exist (standard row names)
    match_fallback = re.search(r"^(INTRO|[SQD]\d+)", sawtooth_id, re.IGNORECASE)
    if match_fallback:
        return match_fallback.group(1).upper()
        
    return "GLOBAL"

# --- Streamlit UI Configuration ---
st.sidebar.header("⚙️ Configuration Matrix")
target_column_input = st.sidebar.text_input("Target Language Column Name", value="Target: philippines")

uploaded_excel = st.file_uploader("1. Upload Sawtooth Export (Excel or CSV)", type=["xlsx", "csv"])
uploaded_docx = st.file_uploader("2. Upload Questionnaire Word File (.docx)", type=["docx"])

if uploaded_excel and uploaded_docx:
    if st.button("🚀 Match and Map Translations", use_container_width=True):
        with st.spinner("Processing strict block verification logic..."):
            try:
                if uploaded_excel.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_excel)
                else:
                    df = pd.read_excel(uploaded_excel)
                
                if 'Source: en' not in df.columns or 'Id' not in df.columns:
                    st.error("Error: Missing required columns 'Id' or 'Source: en'.")
                    st.stop()
                
                # Setup target column structure dynamically
                df[target_column_input] = ""
                df[target_column_input] = df[target_column_input].astype(object)
                df['Source: en'] = df['Source: en'].astype(str)

                # Step 1: Divide the Word file into isolated question blocks
                word_sections = parse_word_by_question_blocks(uploaded_docx)
                translations_added = 0
                
                # Step 2 & 3: Iterate rows and match based on strict parameters
                for idx, row in df.iterrows():
                    source_text = str(row['Source: en']).strip()
                    sawtooth_id = row['Id']
                    
                    if not source_text or source_text.lower() in ['nan', ''] or source_text.isdigit():
                        continue
                    
                    # Identify the Question Code from the Id column
                    q_code = identify_question_code(sawtooth_id)
                    
                    # Fetch lines belonging exclusively to that Question section
                    section_lines = word_sections.get(q_code, word_sections["GLOBAL"])
                    
                    # Look up exact token translation inside that specific block
                    translation = extract_strict_translation(source_text, section_lines)
                    
                    # General backup pass across the entire doc if structural naming didn't match perfectly
                    if not translation and q_code != "GLOBAL":
                        translation = extract_strict_translation(source_text, word_sections["GLOBAL"])
                        
                    if translation:
                        df.at[idx, target_column_input] = translation
                        translations_added += 1

                st.success(f"Successfully processed! Populated {translations_added} accurate translation pairs.")
                
                # Build download buffer
                output = io.BytesIO()
                if uploaded_excel.name.endswith('.csv'):
                    df.to_csv(output, index=False)
                    mime_type = "text/csv"
                    out_filename = "Sawtooth_Strict_Final.csv"
                else:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    out_filename = "Sawtooth_Strict_Final.xlsx"
                
                output.seek(0)
                
                st.download_button(
                    label=f"📥 Download Updated {target_column_input} File",
                    data=output,
                    file_name=out_filename,
                    mime=mime_type,
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"An error occurred during final processing: {str(e)}")
import streamlit as st
import pandas as pd
from docx import Document
from rapidfuzz import fuzz
import io
import re

st.set_page_config(page_title="Strict Sawtooth Layout Automator", layout="centered")

st.title("🎯 Strict Sawtooth Translation Automator")
st.write("Maps translations by verifying BOTH the Question block ID and matching the exact 'Source: en' token.")

# --- Helper Logic ---
def parse_word_sections(docx_file):
    """
    Parses the questionnaire into separate question dictionaries.
    Example: sections['INTRO'] = ["Yes, Oo", "No, Hindi"]
    """
    doc = Document(docx_file)
    sections = {"GLOBAL": []}
    current_section = "GLOBAL"
    
    # Pool all elements sequentially (paragraphs and tables)
    all_lines = []
    for para in doc.paragraphs:
        txt = para.text.strip()
        if txt:
            all_lines.append(txt)
    for table in doc.tables:
        for row in table.rows:
            cells_txt = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells_txt:
                all_lines.append(" | ".join(cells_txt))
                
    # Regex to detect question headers like Intro., S1., Q2., [S1]
    q_pattern = re.compile(r'^(?:\[)?([SQD\d]+[a-zA-Z0-9]*)(?:\])?[\.\s\:]', re.IGNORECASE)
    
    for line in all_lines:
        match = q_pattern.match(line)
        if match:
            # e.g., INTRO, S1, S2
            current_section = match.group(1).upper()
            if current_section not in sections:
                sections[current_section] = []
        
        sections[current_section].append(line)
        sections["GLOBAL"].append(line)
        
    return sections

def find_exact_token_translation(source_text, section_lines):
    """
    Looks inside the isolated question lines for an explicit token split
    where the English side matches our 'Source: en' column text.
    """
    src_clean = source_text.strip().lower()
    
    for line in section_lines:
        parts = []
        if '|' in line:
            parts = [p.strip() for p in line.split('|')]
        elif ',' in line:
            parts = [p.strip() for p in line.split(',')]
        else:
            continue
            
        # The line must have at least an English token and a Translation token
        if len(parts) >= 2:
            first_part_clean = parts[0].strip().lower()
            
            # Check if the English text matches your source cell token exactly (or near exact)
            if first_part_clean == src_clean or fuzz.ratio(first_part_clean, src_clean) > 95:
                translation_candidate = parts[1].strip()
                
                # Filter out numbers if they represent survey data codes
                if translation_candidate.isdigit() and len(parts) > 2:
                    return parts[2].strip()
                    
                return translation_candidate
                
    return None

def extract_clean_q_code(sawtooth_id):
    """
    Extracts base tracking code out of Sawtooth ID strings.
    List 'IntroList' -> INTRO
    List 'S1List' -> S1
    """
    sawtooth_id = str(sawtooth_id)
    match = re.search(r"'([a-zA-Z0-9]+)", sawtooth_id)
    if match:
        # Strip out standard suffixes like 'List' to get back to core question tag
        code = match.group(1)
        code = re.sub(r'(List|Term|Qnr)$', '', code, flags=re.IGNORECASE)
        return code.upper()
    return "GLOBAL"

# --- Frontend Streamlit Configuration Layout ---
st.sidebar.header("⚙️ Configuration Controls")
target_col_name = st.sidebar.text_input("Target Language Column Name", value="Target: philippines")

uploaded_excel = st.file_uploader("1. Upload Sawtooth Export (Excel or CSV)", type=["xlsx", "csv"])
uploaded_docx = st.file_uploader("2. Upload Questionnaire Word File (.docx)", type=["docx"])

if uploaded_excel and uploaded_docx:
    if st.button("🚀 Match Source & Build Sheet", use_container_width=True):
        with st.spinner("Executing strict source-token checks..."):
            try:
                if uploaded_excel.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_excel)
                else:
                    df = pd.read_excel(uploaded_excel)
                
                if 'Source: en' not in df.columns or 'Id' not in df.columns:
                    st.error("Error: The layout sheet must contain 'Id' and 'Source: en' columns.")
                    st.stop()
                
                # Initialize custom targeted translation column
                df[target_col_name] = ""
                df[target_col_name] = df[target_col_name].astype(object)
                df['Source: en'] = df['Source: en'].astype(str)

                # Parse Word document structural matrices
                word_sections = parse_word_sections(uploaded_docx)
                translations_added = 0
                
                for idx, row in df.iterrows():
                    source_text = str(row['Source: en']).strip()
                    sawtooth_id = row['Id']
                    
                    if not source_text or source_text.lower() in ['nan', ''] or source_text.isdigit():
                        continue
                    
                    # 1. Identify context block name (e.g., INTRO)
                    q_code = extract_clean_q_code(sawtooth_id)
                    section_lines = word_sections.get(q_code, word_sections["GLOBAL"])
                    
                    # 2. Strict look-up for English token match
                    translation = find_exact_token_translation(source_text, section_lines)
                    
                    # Global backup pass if code naming variations missed the local block
                    if not translation and q_code != "GLOBAL":
                        translation = find_exact_token_translation(source_text, word_sections["GLOBAL"])
                        
                    if translation:
                        df.at[idx, target_col_name] = translation
                        translations_added += 1

                st.success(f"Success! Strictly matched and populated {translations_added} clean source translations.")
                
                # Build deliverable compilation download buffer
                output = io.BytesIO()
                if uploaded_excel.name.endswith('.csv'):
                    df.to_csv(output, index=False)
                    mime_type = "text/csv"
                    out_filename = "Sawtooth_Strict_Translations.csv"
                else:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    out_filename = "Sawtooth_Strict_Translations.xlsx"
                
                output.seek(0)
                
                st.download_button(
                    label=f"📥 Download Corrected {target_col_name} Sheet",
                    data=output,
                    file_name=out_filename,
                    mime=mime_type,
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"Processing error: {str(e)}")
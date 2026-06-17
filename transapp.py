import streamlit as st
import pandas as pd
from docx import Document
from rapidfuzz import process, fuzz
import io
import re

st.set_page_config(page_title="Sawtooth Translation Automator", layout="centered")

st.title("🔄 Clean Sawtooth Translation Automator")
st.write("Extracts ONLY the target language translations. Automated instruction & English text filtering.")

# --- Helper Functions ---
def clean_instruction_text(text):
    """Removes standard survey routing or structural instructions."""
    if not text:
        return ""
    # Strip bracketed elements e.g. [Select one], [Exclusive]
    text = re.sub(r'\[.*?\]', '', text)
    # Strip common routing keywords
    routing_keywords = [
        r'\bask\s+all\b', r'\bterminate\b', r'\bcontinue\b', 
        r'\bshow\s+on\s+\d+\s+page\b', r'\bsingle\s+select\b', r'\bmulti\s+select\b'
    ]
    for pattern in routing_keywords:
        text = re.sub(pattern, '', text, flags=re.IGNORECASE)
    
    # Clean up trailing pipes, commas, or multiple whitespaces left behind
    text = re.sub(r'[\s,|]+$', '', text)
    text = re.sub(r'^[\s,|]+', '', text)
    return text.strip()

def parse_word_by_sections(docx_file):
    """Parses Word document sequentially into isolated Question Blocks."""
    doc = Document(docx_file)
    sections = {"GLOBAL": []}
    current_section = "GLOBAL"
    
    all_lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            all_lines.append(text)
    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_text:
                all_lines.append(" | ".join(row_text))
                
    q_pattern = re.compile(r'^(?:\[)?([SQD]\d+[a-zA-Z0-9]*)(?:\])?[\.\s\:]', re.IGNORECASE)
    
    for line in all_lines:
        match = q_pattern.match(line)
        if match:
            current_section = match.group(1).upper()
            sections[current_section] = []
        
        sections[current_section].append(line)
        sections["GLOBAL"].append(line)
        
    return sections

def find_pure_translation(source_text, section_lines):
    """Finds the source text in a row and extracts ONLY the translation part."""
    source_clean = source_text.strip().lower()
    
    for line in section_lines:
        # Determine the delimiter used in this block (comma or table pipe)
        parts = []
        if '|' in line:
            parts = [p.strip() for p in line.split('|')]
        elif ',' in line:
            parts = [p.strip() for p in line.split(',')]
            
        if len(parts) >= 2:
            # Look for which index matches your English source text
            for idx, part in enumerate(parts):
                part_clean = part.lower().strip()
                if part_clean == source_clean or fuzz.ratio(part_clean, source_clean) > 92:
                    # The next part is usually the translation (e.g., Parts: [English, Translated, Code])
                    if idx + 1 < len(parts):
                        possible_translation = parts[idx + 1]
                        # If the next part is just a data code number, look backward or look ahead
                        if possible_translation.isdigit() and idx > 0:
                            possible_translation = parts[idx - 1]
                        
                        cleaned_trans = clean_instruction_text(possible_translation)
                        if cleaned_trans and cleaned_trans.lower() != source_clean:
                            return cleaned_trans

    # Fallback to sequential paragraph lines (English paragraph followed by Translation paragraph)
    for i in range(len(section_lines) - 1):
        line_clean = section_lines[i].strip().lower()
        if line_clean == source_clean or fuzz.ratio(line_clean, source_clean) > 88:
            next_line = section_lines[i+1].strip()
            cleaned_trans = clean_instruction_text(next_line)
            if cleaned_trans and cleaned_trans.lower() != line_clean:
                return cleaned_trans
                
    return None

def extract_q_code_from_id(sawtooth_id):
    """Extracts question keys from Sawtooth ID string."""
    sawtooth_id = str(sawtooth_id)
    match = re.search(r"'(S|Q|D)(\d+[a-zA-Z0-9]*)", sawtooth_id, re.IGNORECASE)
    if match:
        return (match.group(1) + match.group(2)).upper()
    match_fallback = re.search(r"'([^']+)'", sawtooth_id)
    if match_fallback:
        return re.sub(r'(List|Term|Qnr)$', '', match_fallback.group(1), flags=re.IGNORECASE).upper()
    return "GLOBAL"

# --- Streamlit File Upload System ---
uploaded_excel = st.file_uploader("1. Upload Sawtooth Export (Excel or CSV)", type=["xlsx", "csv"])
uploaded_docx = st.file_uploader("2. Upload Translated Word Doc (.docx)", type=["docx"])

if uploaded_excel and uploaded_docx:
    if st.button("🚀 Process Pure Translations", use_container_width=True):
        with st.spinner("Isolating pure translations and cleaning instructions..."):
            try:
                if uploaded_excel.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_excel)
                else:
                    df = pd.read_excel(uploaded_excel)
                
                if 'Source: en' not in df.columns or 'Id' not in df.columns:
                    st.error("Error: Missing required columns 'Id' or 'Source: en'.")
                    st.stop()
                
                if 'Target: en' not in df.columns:
                    df['Target: en'] = ""
                    
                df['Target: en'] = df['Target: en'].astype(object)
                df['Source: en'] = df['Source: en'].astype(str)

                # Map Word doc structure
                word_sections = parse_word_by_sections(uploaded_docx)
                translations_added = 0
                
                for idx, row in df.iterrows():
                    source_text = str(row['Source: en']).strip()
                    sawtooth_id = row['Id']
                    
                    if not source_text or source_text.lower() in ['nan', ''] or source_text.isdigit():
                        continue
                    
                    q_code = extract_q_code_from_id(sawtooth_id)
                    section_lines = word_sections.get(q_code, word_sections["GLOBAL"])
                    
                    # Core matching logic targeting isolated elements
                    translation = find_pure_translation(source_text, section_lines)
                    
                    if not translation and q_code != "GLOBAL":
                        translation = find_pure_translation(source_text, word_sections["GLOBAL"])
                        
                    if translation:
                        df.at[idx, 'Target: en'] = translation
                        translations_added += 1

                st.success(f"Done! Cleaned and populated {translations_added} target elements.")
                
                # Output Generator
                output = io.BytesIO()
                if uploaded_excel.name.endswith('.csv'):
                    df.to_csv(output, index=False)
                    mime_type = "text/csv"
                    out_filename = "Sawtooth_Pure_Translations.csv"
                else:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    out_filename = "Sawtooth_Pure_Translations.xlsx"
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Cleaned Sawtooth Sheet",
                    data=output,
                    file_name=out_filename,
                    mime=mime_type,
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
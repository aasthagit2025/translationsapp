import streamlit as st
import pandas as pd
from docx import Document
from rapidfuzz import process, fuzz
import io
import re

st.set_page_config(page_title="Sawtooth Translation Automator", layout="centered")

st.title("🔄 Context-Aware Sawtooth Translation Automator")
st.write("Populates translation columns by targeting the specific question blocks defined in the Excel ID column.")

# --- Helper Functions ---
def parse_word_by_sections(docx_file):
    """
    Parses the Word document into a dictionary of question sections.
    Key: Question code extracted from text headers (e.g., 'S1', 'Q5')
    Value: List of text strings found exclusively within that question block.
    """
    doc = Document(docx_file)
    sections = {"GLOBAL": []}
    current_section = "GLOBAL"
    
    # Pool all paragraphs and tables sequentially
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
                
    # Regex to detect lines initiating question boundaries: S1., Q2., S2a., Q14b:, [S1]
    q_pattern = re.compile(r'^(?:\[)?([SQD]\d+[a-zA-Z0-9]*)(?:\])?[\.\s\:]', re.IGNORECASE)
    
    for line in all_lines:
        match = q_pattern.match(line)
        if match:
            current_section = match.group(1).upper()
            sections[current_section] = []
        
        sections[current_section].append(line)
        sections["GLOBAL"].append(line) # Back-up copy for universal matches
        
    return sections

def find_translation_in_section(source_text, section_lines):
    """
    Searches for the corresponding target translation string within isolated section lines.
    Handles comma-separated list entries and consecutive paragraph rows.
    """
    source_clean = source_text.strip().lower()
    
    # 1. First Pass: Check for explicit choice lists (e.g., "The Philippines,Pilipinas,4")
    for line in section_lines:
        if ',' in line:
            parts = [p.strip() for p in line.split(',')]
            if len(parts) >= 2:
                if parts[0].strip().lower() == source_clean:
                    return parts[1]
                if fuzz.ratio(parts[0].strip().lower(), source_clean) > 92:
                    return parts[1]

    # 2. Second Pass: Check consecutive lines (English text row followed immediately by Tagalog)
    for i in range(len(section_lines) - 1):
        line_clean = section_lines[i].strip().lower()
        
        # Exact matching or highly accurate fuzzy alignment
        if source_clean in line_clean or fuzz.ratio(line_clean, source_clean) > 85:
            next_line = section_lines[i+1].strip()
            # Eliminate structural tracking labels or instruction lines like [Select one]
            if next_line and not next_line.startswith('[') and next_line.lower() != line_clean:
                return next_line
                
    return None

def extract_q_code_from_id(sawtooth_id):
    """
    Extracts the pure base question token from the Sawtooth structural Id string.
    e.g., "List 'S1List' - Item 1 - Display Text" -> "S1"
    e.g., "Question 'S5xMYTerm' - Body Text" -> "S5XMYTERM"
    """
    sawtooth_id = str(sawtooth_id)
    # Match alphanumeric sequences starting with S, Q, or D inside single quotes
    match = re.search(r"'(S|Q|D)(\d+[a-zA-Z0-9]*)", sawtooth_id, re.IGNORECASE)
    if match:
        return (match.group(1) + match.group(2)).upper()
    
    # Fallback cleanup for specialized IDs
    match_fallback = re.search(r"'([^']+)'", sawtooth_id)
    if match_fallback:
        clean_code = re.sub(r'(List|Term|Qnr)$', '', match_fallback.group(1), flags=re.IGNORECASE)
        return clean_code.upper()
        
    return "GLOBAL"

# --- UI File Uploads ---
uploaded_excel = st.file_uploader("1. Upload Sawtooth Export (Excel or CSV)", type=["xlsx", "csv"])
uploaded_docx = st.file_uploader("2. Upload Translated Word Doc (.docx)", type=["docx"])

if uploaded_excel and uploaded_docx:
    if st.button("🚀 Process Context-Aware Translations", use_container_width=True):
        with st.spinner("Analyzing structural hierarchies and assigning fields..."):
            try:
                # Load structural configuration sheet
                if uploaded_excel.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_excel)
                else:
                    df = pd.read_excel(uploaded_excel)
                
                if 'Source: en' not in df.columns or 'Id' not in df.columns:
                    st.error("Error: Missing structural column coordinates 'Id' or 'Source: en'.")
                    st.stop()
                
                if 'Target: en' not in df.columns:
                    df['Target: en'] = ""
                    
                df['Target: en'] = df['Target: en'].astype(object)
                df['Source: en'] = df['Source: en'].astype(str)

                # Step 1: Divide the Word questionnaire context into structured blocks
                word_sections = parse_word_by_sections(uploaded_docx)

                translations_added = 0
                
                # Step 2: Correlate rows with context verification
                for idx, row in df.iterrows():
                    source_text = str(row['Source: en']).strip()
                    sawtooth_id = row['Id']
                    
                    if not source_text or source_text.lower() in ['nan', ''] or source_text.isdigit():
                        continue
                    
                    # Compute question label context (e.g. "S1")
                    q_code = extract_q_code_from_id(sawtooth_id)
                    
                    # Pull lines specifically tied to that question boundary
                    section_lines = word_sections.get(q_code, word_sections["GLOBAL"])
                    
                    # Look up translation matching in that isolated section
                    translation = find_translation_in_section(source_text, section_lines)
                    
                    # Universal Fallback if no matching text was isolated inside that specific section block
                    if not translation and q_code != "GLOBAL":
                        translation = find_translation_in_section(source_text, word_sections["GLOBAL"])
                        
                    if translation:
                        df.at[idx, 'Target: en'] = translation
                        translations_added += 1

                st.success(f"Done! Safely matched and populated {translations_added} target fields.")
                
                # Prepare compilation buffer for delivery
                output = io.BytesIO()
                if uploaded_excel.name.endswith('.csv'):
                    df.to_csv(output, index=False)
                    mime_type = "text/csv"
                    out_filename = "Sawtooth_Translations_Targeted.csv"
                else:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    out_filename = "Sawtooth_Translations_Targeted.xlsx"
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Corrected Sawtooth Sheet",
                    data=output,
                    file_name=out_filename,
                    mime=mime_type,
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
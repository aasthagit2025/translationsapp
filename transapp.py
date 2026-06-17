import streamlit as st
import pandas as pd
from docx import Document
from rapidfuzz import process, fuzz
import io

st.set_page_config(page_title="Sawtooth Translation Automator", layout="centered")

st.title("🔄 Sawtooth Translation Automator")
st.write("Upload your English Sawtooth Excel export and your translated Word Questionnaire to auto-populate the translation columns.")

# --- Helper Functions ---
def extract_text_from_word(docx_file):
    """Extracts all text blocks from paragraphs and tables in the uploaded Word file."""
    doc = Document(docx_file)
    raw_blocks = []
    
    # Extract from paragraphs
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            raw_blocks.append(text)
            
    # Extract from tables
    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_text:
                raw_blocks.append(" | ".join(row_text))
                
    return raw_blocks

def build_translation_mappings(raw_blocks):
    """Parses text lines to capture comma-separated list options or consecutive translated rows."""
    mapping_dict = {}
    
    for i in range(len(raw_blocks)):
        current_line = raw_blocks[i]
        
        # Scenario A: Comma separated inline list items (e.g., 'The Philippines,Pilipinas,4')
        if ',' in current_line:
            parts = [p.strip() for p in current_line.split(',')]
            if len(parts) >= 2:
                eng_opt = parts[0]
                trans_opt = parts[1]
                if eng_opt and not eng_opt.isdigit() and trans_opt:
                    mapping_dict[eng_opt] = trans_opt
        
        # Scenario B: Pairs of sequential paragraphs (English line followed by Translated line)
        if i < len(raw_blocks) - 1:
            next_line = raw_blocks[i+1]
            if current_line != next_line and not current_line.startswith('[') and not next_line.startswith('['):
                mapping_dict[current_line] = next_line
                
    return mapping_dict

# --- UI File Uploads ---
uploaded_excel = st.file_uploader("1. Upload Sawtooth Export (Excel or CSV)", type=["xlsx", "csv"])
uploaded_docx = st.file_uploader("2. Upload Translated Word Doc (.docx)", type=["docx"])

if uploaded_excel and uploaded_docx:
    if st.button("🚀 Process & Generate Translations", use_container_width=True):
        with st.spinner("Analyzing text alignment & running fuzzy matching..."):
            try:
                # Load Excel/CSV
                if uploaded_excel.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_excel)
                else:
                    df = pd.read_excel(uploaded_excel)
                
                # Check column requirements
                if 'Source: en' not in df.columns or 'Target: en' not in df.columns:
                    st.error("Error: The uploaded sheet must contain 'Source: en' and 'Target: en' columns.")
                    st.stop()

                # Process Word Document
                word_blocks = extract_text_from_word(uploaded_docx)
                translation_lookup = build_translation_mappings(word_blocks)
                english_keys = list(translation_lookup.keys())

                # Process Matching
                translations_added = 0
                for idx, row in df.iterrows():
                    source_text = str(row['Source: en']).strip() if pd.notnull(row['Source: en']) else ""
                    
                    if not source_text or source_text.isdigit():
                        continue
                        
                    # 1. Direct exact match check
                    if source_text in translation_lookup:
                        df.at[idx, 'Target: en'] = translation_lookup[source_text]
                        translations_added += 1
                    else:
                        # 2. Fuzzy fallback match (85% threshold)
                        match = process.extractOne(source_text, english_keys, scorer=fuzz.token_sort_ratio)
                        if match and match[1] >= 85:
                            matched_eng_key = match[0]
                            df.at[idx, 'Target: en'] = translation_lookup[matched_eng_key]
                            translations_added += 1

                st.success(f"Done! Automatically populated {translations_added} translation strings.")
                
                # Prepare binary output buffer for user download
                output = io.BytesIO()
                if uploaded_excel.name.endswith('.csv'):
                    df.to_csv(output, index=False)
                    mime_type = "text/csv"
                    out_filename = "Sawtooth_Translations_Ready.csv"
                else:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    out_filename = "Sawtooth_Translations_Ready.xlsx"
                
                output.seek(0)
                
                # Download Button
                st.download_button(
                    label="📥 Download Updated Sawtooth Sheet",
                    data=output,
                    file_name=out_filename,
                    mime=mime_type,
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
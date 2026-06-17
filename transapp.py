import streamlit as st
import pandas as pd
from docx import Document
from rapidfuzz import process, fuzz
import io
import re

st.set_page_config(page_title="Sawtooth Dual-Filter Automator", layout="centered")

st.title("🎨🔄 Sawtooth Language & Color Automator")
st.write("Extracts clean translations using an interactive target language column assignment and a Word font color filter.")

# --- Helper Functions ---
def is_color_match(run, target_color):
    """Checks if a Word document run text segment matches the user's targeted color choice."""
    font_color = run.font.color
    if not font_color or font_color.rgb is None:
        return False
        
    hex_str = str(font_color.rgb).lower()
    tc = target_color.lower().strip()
    
    if tc == 'green':
        # Detect standard office green variations (e.g., '00b050', '008000')
        if hex_str.startswith('00b050') or hex_str.startswith('0080') or hex_str.startswith('4c9') or hex_str.startswith('00a'):
            return True
        # Algorithmic fallback: Check if Green channel dominates Red and Blue
        if len(hex_str) == 6:
            try:
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                if g > r * 1.2 and g > b * 1.2 and g > 40:
                    return True
            except ValueError:
                pass
    elif tc == 'red' and (hex_str.startswith('ff0000') or hex_str.startswith('c00000')):
        return True
    elif tc == 'blue' and (hex_str.startswith('0000ff') or hex_str.startswith('1f4e79')):
        return True
    elif tc.replace('#', '') in hex_str:
        return True
        
    return False

def extract_colored_text_by_sections(docx_file, target_color):
    """Parses Word file chronologically, extracting ONLY elements matching the specified font color."""
    doc = Document(docx_file)
    sections = {"GLOBAL": []}
    current_section = "GLOBAL"
    
    q_pattern = re.compile(r'^(?:\[)?([SQD]\d+[a-zA-Z0-9]*)(?:\])?[\.\s\:]', re.IGNORECASE)
    
    # Process Paragraphs
    for para in doc.paragraphs:
        plain_text = para.text.strip()
        match = q_pattern.match(plain_text)
        if match:
            current_section = match.group(1).upper()
            if current_section not in sections:
                sections[current_section] = []
                
        colored_parts = [run.text.strip() for run in para.runs if is_color_match(run, target_color) and run.text.strip()]
        if colored_parts:
            combined = " ".join(colored_parts)
            sections[current_section].append(combined)
            sections["GLOBAL"].append(combined)
            
    # Process Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p_text = p.text.strip()
                    match = q_pattern.match(p_text)
                    if match:
                        current_section = match.group(1).upper()
                        if current_section not in sections:
                            sections[current_section] = []
                            
                    cell_colored_parts = [run.text.strip() for run in p.runs if is_color_match(run, target_color) and run.text.strip()]
                    if cell_colored_parts:
                        combined = " ".join(cell_colored_parts)
                        sections[current_section].append(combined)
                        sections["GLOBAL"].append(combined)
                        
    return sections

def extract_q_code_from_id(sawtooth_id):
    """Extracts structural question prefix context values from Sawtooth ID string."""
    sawtooth_id = str(sawtooth_id)
    match = re.search(r"'(S|Q|D)(\d+[a-zA-Z0-9]*)", sawtooth_id, re.IGNORECASE)
    if match:
        return (match.group(1) + match.group(2)).upper()
    match_fallback = re.search(r"'([^']+)'", sawtooth_id)
    if match_fallback:
        return re.sub(r'(List|Term|Qnr)$', '', match_fallback.group(1), flags=re.IGNORECASE).upper()
    return "GLOBAL"

# --- Streamlit Sidebar Config User Controls ---
st.sidebar.header("⚙️ Configuration Controls")
target_col_input = st.sidebar.text_input("Target Language Column Name", value="Target: en")
color_dropdown = st.sidebar.selectbox("Translation Font Color in Word", ["Green", "Red", "Blue", "Custom Hex Code"])

if color_dropdown == "Custom Hex Code":
    selected_color = st.sidebar.text_input("Enter exact Font Hex Code (e.g., 00B050)", value="00B050")
else:
    selected_color = color_dropdown

# --- File Uploader Frontend UI ---
uploaded_excel = st.file_uploader("1. Upload Sawtooth Export Layout Sheet (Excel or CSV)", type=["xlsx", "csv"])
uploaded_docx = st.file_uploader("2. Upload Translated Questionnaire Word File (.docx)", type=["docx"])

if uploaded_excel and uploaded_docx:
    if st.button("🚀 Process Dual-Input Matrix Alignment", use_container_width=True):
        with st.spinner(f"Scraping '{selected_color}' translations and targeting column '{target_col_input}'..."):
            try:
                # Load configuration layouts
                if uploaded_excel.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_excel)
                else:
                    df = pd.read_excel(uploaded_excel)
                
                if 'Source: en' not in df.columns or 'Id' not in df.columns:
                    st.error("Error: The layout sheet must contain structural columns 'Id' and 'Source: en'.")
                    st.stop()
                
                # Check or initialize target column dynamically
                if target_col_input not in df.columns:
                    df[target_col_input] = ""
                    
                df[target_col_input] = df[target_col_input].astype(object)
                df['Source: en'] = df['Source: en'].astype(str)

                # Process color filtering
                word_sections = extract_colored_text_by_sections(uploaded_docx, selected_color)
                translations_added = 0
                
                for idx, row in df.iterrows():
                    source_text = str(row['Source: en']).strip()
                    sawtooth_id = row['Id']
                    
                    if not source_text or source_text.lower() in ['nan', ''] or source_text.isdigit():
                        continue
                        
                    q_code = extract_q_code_from_id(sawtooth_id)
                    colored_pool = word_sections.get(q_code, word_sections["GLOBAL"])
                    
                    if not colored_pool:
                        colored_pool = word_sections["GLOBAL"]
                        
                    if colored_pool:
                        # Direct structural positioning match (e.g., matching item option sequences)
                        item_match = re.search(r"Item (\d+)", str(sawtooth_id))
                        if item_match:
                            item_index = int(item_match.group(1)) - 1
                            if item_index < len(colored_pool):
                                df.at[idx, target_col_input] = colored_pool[item_index]
                                translations_added += 1
                                continue
                        
                        # Fuzzy fallback sequence mapping matching context strings
                        best_match = process.extractOne(source_text, colored_pool, scorer=fuzz.token_sort_ratio)
                        if best_match and best_match[1] >= 35:
                            df.at[idx, target_col_input] = best_match[0]
                            translations_added += 1

                st.success(f"Success! Extracted and mapped {translations_added} clean localized target rows.")
                
                # Create deployment delivery buffer
                output = io.BytesIO()
                if uploaded_excel.name.endswith('.csv'):
                    df.to_csv(output, index=False)
                    mime_type = "text/csv"
                    out_filename = f"Sawtooth_Translations_{target_col_input.replace(':', '')}.csv"
                else:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    out_filename = f"Sawtooth_Translations_{target_col_input.replace(':', '')}.xlsx"
                
                output.seek(0)
                
                st.download_button(
                    label=f"📥 Download Updated {target_col_input} Export File",
                    data=output,
                    file_name=out_filename,
                    mime=mime_type,
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"An error occurred during multi-input processing: {str(e)}")
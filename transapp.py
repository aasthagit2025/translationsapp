import streamlit as st
import pandas as pd
from docx import Document
from rapidfuzz import fuzz
import io
import re

st.set_page_config(page_title="Color-Strict Sawtooth Layout Automator", layout="centered")

st.title("🎯 Color-Strict Sawtooth Translation Automator")
st.write("Extracts ONLY the text matching your designated Font Color within the targeted Question ID section.")

# --- Core Color-Strict Matrix Logic ---
def is_run_color_match(run, target_color):
    """Verifies if an individual word segment matches the target color description."""
    font_color = run.font.color
    if not font_color or font_color.rgb is None:
        return False
        
    hex_str = str(font_color.rgb).lower()
    tc = target_color.lower().strip()
    
    if tc == 'green':
        # Target standard Office Green variations found in Qre docs (e.g., '00b050', '008000')
        if hex_str.startswith('00b050') or hex_str.startswith('0080') or hex_str.startswith('4c9') or hex_str.startswith('00a'):
            return True
        # Algorithmic check: Is Green significantly stronger than Red and Blue channels?
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

def parse_word_sections_by_color(docx_file, target_color):
    """
    Chronologically reads the questionnaire and maps text tokens.
    Separates items into lists of structured segments matching the target color inside each Question block.
    """
    doc = Document(docx_file)
    sections = {"GLOBAL": []}
    current_section = "GLOBAL"
    
    # Structural rule to spot new questions (e.g., Intro., S1., Q5.)
    q_pattern = re.compile(r'^(?:\[)?(INTRO|[SQD]\d+[a-zA-Z0-9]*)(?:\])?[\.\s\:]', re.IGNORECASE)
    
    # 1. Process Paragraph rows
    for para in doc.paragraphs:
        plain_text = para.text.strip()
        match = q_pattern.match(plain_text)
        if match:
            current_section = match.group(1).upper()
            if current_section not in sections:
                sections[current_section] = []
        
        # Analyze individual delimited segments in a line
        raw_parts = [p.strip() for p in plain_text.split(',') if p.strip()] if ',' in plain_text else [plain_text]
        
        # Cross reference paragraph chunks with font formatting runs
        for run in para.runs:
            run_text = run.text.strip()
            if run_text and is_run_color_match(run, target_color):
                # If it's isolated colored target text, append to section tracking
                sections[current_section].append((plain_text, run_text))
                sections["GLOBAL"].append((plain_text, run_text))
                
    # 2. Process Table cells
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
                            
                    for run in p.runs:
                        run_text = run.text.strip()
                        if run_text and is_run_color_match(run, target_color):
                            sections[current_section].append((p_text, run_text))
                            sections["GLOBAL"].append((p_text, run_text))
                            
    return sections

def get_clean_question_id(sawtooth_id):
    """Isolates the target Question token from Sawtooth structure naming formats."""
    sawtooth_id = str(sawtooth_id)
    match = re.search(r"'([a-zA-Z0-9]+)", sawtooth_id)
    if match:
        code = match.group(1)
        # Strip trailing functional suffixes like 'List' to get back to core question tag
        code = re.sub(r'(List|Term|Qnr|Labels)$', '', code, flags=re.IGNORECASE)
        return code.upper()
    return "GLOBAL"

def find_colored_translation(source_text, section_tuples):
    """
    Finds the exact line match inside the section pool and extracts 
    ONLY the part that matches the designated color run.
    """
    src_clean = source_text.strip().lower()
    
    # Step 1: Strict Line Match via row data splits
    for plain_line, color_text in section_tuples:
        # Check if the overall row text contains our English token
        parts = [p.strip().lower() for p in plain_line.split(',') if p.strip()] if ',' in plain_line else [plain_line.lower()]
        
        if src_clean in parts or any(fuzz.ratio(p, src_clean) > 96 for p in parts):
            # Avoid re-adding English source if it was formatted in the color run
            if color_text.lower() != src_clean and not color_text.isdigit():
                return color_text
                
    return None

# --- UI Layout Controls ---
st.sidebar.header("⚙️ Target Parameter Matrix")
target_column_input = st.sidebar.text_input("Target Language Column Name", value="Target: philippines")
color_selection = st.sidebar.selectbox("Translation Font Color in Word", ["Green", "Red", "Blue", "Custom Hex"])

if color_selection == "Custom Hex":
    chosen_color = st.sidebar.text_input("Enter Font Hex (e.g., 00B050)", value="00B050")
else:
    chosen_color = color_selection

uploaded_excel = st.file_uploader("1. Upload Sawtooth Export (Excel or CSV)", type=["xlsx", "csv"])
uploaded_docx = st.file_uploader("2. Upload Questionnaire Word File (.docx)", type=["docx"])

if uploaded_excel and uploaded_docx:
    if st.button("🚀 Run Color-Strict Extraction", use_container_width=True):
        with st.spinner(f"Scraping pure '{chosen_color}' tokens for target '{target_column_input}'..."):
            try:
                if uploaded_excel.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_excel)
                else:
                    df = pd.read_excel(uploaded_excel)
                
                if 'Source: en' not in df.columns or 'Id' not in df.columns:
                    st.error("Error: Missing critical layout coordinates 'Id' or 'Source: en'.")
                    st.stop()
                    
                df[target_column_input] = ""
                df[target_column_input] = df[target_column_input].astype(object)
                df['Source: en'] = df['Source: en'].astype(str)

                # Map Word doc structure isolating targeted color run elements
                word_color_sections = parse_word_sections_by_color(uploaded_docx, chosen_color)
                translations_added = 0
                
                for idx, row in df.iterrows():
                    source_text = str(row['Source: en']).strip()
                    sawtooth_id = row['Id']
                    
                    if not source_text or source_text.lower() in ['nan', ''] or source_text.isdigit():
                        continue
                        
                    # 1. Check ID Column to find Question No
                    q_code = get_clean_question_id(sawtooth_id)
                    section_tuples = word_color_sections.get(q_code, word_color_sections["GLOBAL"])
                    
                    # 2. Check Source Column text matching inside isolated color segments
                    translation = find_colored_translation(source_text, section_tuples)
                    
                    # Backup global sweep if naming conventions missed the section dictionary box
                    if not translation and q_code != "GLOBAL":
                        translation = find_colored_translation(source_text, word_color_sections["GLOBAL"])
                        
                    if translation:
                        df.at[idx, target_column_input] = translation
                        translations_added += 1

                st.success(f"Success! Populated {translations_added} clean, color-filtered target matches.")
                
                # Output Generator
                output = io.BytesIO()
                if uploaded_excel.name.endswith('.csv'):
                    df.to_csv(output, index=False)
                    mime_type = "text/csv"
                    out_filename = "Sawtooth_ColorStrict_Output.csv"
                else:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    out_filename = "Sawtooth_ColorStrict_Output.xlsx"
                
                output.seek(0)
                
                st.download_button(
                    label=f"📥 Download Finalized {target_column_input} Sheet",
                    data=output,
                    file_name=out_filename,
                    mime=mime_type,
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"Processing error: {str(e)}")
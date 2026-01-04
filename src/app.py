# src/app.py
import streamlit as st
import pandas as pd
from docx import Document
import json

# Import your logic
from mapper import normalize_header, map_row_data

def extract_all_keys(data, keys=None):
    """
    Recursively extract all keys from a JSON structure.
    Returns a flat set of all possible field names.
    """
    if keys is None:
        keys = set()
    
    if isinstance(data, dict):
        for key, value in data.items():
            keys.add(key)
            extract_all_keys(value, keys)
    elif isinstance(data, list) and data:
        # Look at first item in list
        extract_all_keys(data[0], keys)
    
    return keys

st.title("Flexible Exports Pilot")

# 1. File Uploads
template_file = st.file_uploader("Upload Word Template", type="docx")
schema_file = st.file_uploader("Upload JSON Schema", type="json")

if template_file and schema_file:
    # Load data
    schema_data = json.load(schema_file)
    doc = Document(template_file)
    
    # 2. Check if document has tables
    if not doc.tables:
        st.error("⚠️ The uploaded template has no tables. Please upload a template with at least one table.")
        st.stop()
    
    # Let user select which table to use
    st.subheader("Select Table")
    st.write(f"Found {len(doc.tables)} table(s) in the document.")
    
    table_index = st.selectbox(
        "Which table contains the data to map?",
        options=range(len(doc.tables)),
        format_func=lambda i: f"Table {i + 1}"
    )
    
    table = doc.tables[table_index]
    
    # Check if table has rows
    if not table.rows:
        st.error("⚠️ The selected table has no rows.")
        st.stop()
    
    word_headers = [cell.text for cell in table.rows[0].cells]
    
    # 3. Automatic Inference
    # Extract all possible keys from the JSON schema (no assumptions about structure)
    available_json_keys = sorted(list(extract_all_keys(schema_data)))
    
    # Generate initial guesses
    mappings = []
    for header in word_headers:
        normalized = normalize_header(header)
        # Guess the match, or default to None
        match = normalized if normalized in available_json_keys else None
        mappings.append({"Word Column": header, "Mapped JSON Key": match})

    # 4. Manual Review (The "Option to do changes")
    st.subheader("Review Mappings")
    st.write("We automatically matched these columns. Please correct any mistakes.")
    
    # Create an editable DataFrame
    df = pd.DataFrame(mappings)
    
    # Use st.data_editor with a dropdown column config
    edited_df = st.data_editor(
        df,
        column_config={
            "Mapped JSON Key": st.column_config.SelectboxColumn(
                "JSON Field",
                options=available_json_keys + [None],
                required=True
            )
        },
        hide_index=True,
        use_container_width=True
    )

    # 5. Generate Button
    if st.button("Generate Document"):
        # Convert the edited UI table back into a dictionary
        final_mapping = dict(zip(edited_df["Word Column"], edited_df["Mapped JSON Key"]))
        
        # Pass 'final_mapping' to your doc_generator.py function (TODO)
        st.success("Document generated! (Logic to be connected)")
        
        # st.download_button(...) # Allow user to download the result
import streamlit as st
import pandas as pd
from docx import Document
import json
from datetime import datetime
from io import BytesIO
import re

st.set_page_config(page_title="Document Generator", layout="wide")

st.title("📄 Document Generator")
st.markdown("Upload your template, mapping config, and optionally existing data - then fill in or edit the values to generate your document.")

# ============================================================
# File Uploaders
# ============================================================

col1, col2, col3 = st.columns(3)

with col1:
    template_file = st.file_uploader("📄 Word Template", type="docx", help="The Word document template")

with col2:
    mapping_file = st.file_uploader("🗺️ Mapping Config", type="json", help="The mapping configuration from the mapper")

with col3:
    data_file = st.file_uploader("📋 Existing Data (Optional)", type="json", help="Pre-fill with existing JSON data")

if template_file and mapping_file:
    # Load files
    doc = Document(template_file)
    mapping_config = json.load(mapping_file)
    
    # Load existing data if provided
    existing_data = {}
    if data_file:
        existing_data = json.load(data_file)
        st.success(f"✓ Loaded existing data with {len(existing_data)} top-level keys")
    
    st.markdown("---")
    
    # Extract fields from mapping config
    fields_data = []
    for json_path, config in mapping_config.items():
        field_name = config["field_name"]
        field_format = config.get("format")
        group = config.get("group", "ungrouped")
        
        # Get existing value if available
        path_parts = json_path.replace("[]", "").split(".")
        existing_value = existing_data
        for part in path_parts:
            if isinstance(existing_value, dict):
                existing_value = existing_value.get(part, "")
            else:
                existing_value = ""
                break
        
        if not existing_value:
            existing_value = ""
        
        fields_data.append({
            "path": json_path,
            "name": field_name,
            "format": field_format,
            "group": group,
            "value": str(existing_value),
            "doc_location": config.get("document_location", {})
        })
    
    # Group fields by section
    fields_by_group = {}
    for field in fields_data:
        grp = field["group"]
        fields_by_group.setdefault(grp, []).append(field)
    
    # Display form for each group
    st.subheader("📝 Enter/Edit Data")
    st.caption("Fill in the values for each field. Date fields show a calendar picker 📅, email fields validate format 📧")
    
    # Store values in session state
    if "field_values" not in st.session_state:
        st.session_state.field_values = {f["path"]: f["value"] for f in fields_data}
    
    # Create tabs for each group
    group_tabs = st.tabs([f"📁 {grp}" for grp in sorted(fields_by_group.keys())])
    
    for tab, (group_name, fields) in zip(group_tabs, sorted(fields_by_group.items())):
        with tab:
            for field in fields:
                path = field["path"]
                name = field["name"]
                field_format = field["format"]
                current_value = st.session_state.field_values.get(path, "")
                
                # Create input based on format
                if field_format == "date":
                    # Date picker
                    label = f"📅 {path}"
                    try:
                        # Try to parse existing date
                        if current_value:
                            default_date = datetime.strptime(current_value, "%Y-%m-%d").date()
                        else:
                            default_date = datetime.now().date()
                    except:
                        default_date = datetime.now().date()
                    
                    date_value = st.date_input(
                        label,
                        value=default_date,
                        key=f"date_{path}"
                    )
                    st.session_state.field_values[path] = date_value.strftime("%Y-%m-%d")
                
                elif field_format == "email":
                    # Email input with validation
                    label = f"📧 {path}"
                    email_value = st.text_input(
                        label,
                        value=current_value,
                        key=f"email_{path}",
                        placeholder="user@example.com"
                    )
                    
                    # Validate email
                    if email_value:
                        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                        if not re.match(email_pattern, email_value):
                            st.warning("⚠️ Invalid email format")
                    
                    st.session_state.field_values[path] = email_value
                
                else:
                    # Regular text input
                    text_value = st.text_input(
                        path,
                        value=current_value,
                        key=f"text_{path}"
                    )
                    st.session_state.field_values[path] = text_value
    
    # Show data preview
    st.markdown("---")
    
    with st.expander("🔍 Preview JSON Data"):
        # Build the nested JSON structure
        output_json = {}
        for path, value in st.session_state.field_values.items():
            if value:  # Only include non-empty values
                parts = path.replace("[]", "").split(".")
                current = output_json
                for part in parts[:-1]:
                    if part not in current:
                        current[part] = {}
                    current = current[part]
                current[parts[-1]] = value
        
        st.json(output_json)
    
    # Generate document
    st.markdown("---")
    col1, col2 = st.columns([1, 3])
    
    with col1:
        if st.button("🚀 Generate Document", type="primary"):
            # Build the nested JSON
            output_json = {}
            for path, value in st.session_state.field_values.items():
                if value:
                    parts = path.replace("[]", "").split(".")
                    current = output_json
                    for part in parts[:-1]:
                        if part not in current:
                            current[part] = {}
                        current = current[part]
                    current[parts[-1]] = value
            
            # Apply values to document
            # Parse document to get all text segments (reusing logic from mapper)
            from docx.oxml.table import CT_Tbl
            from docx.oxml.text.paragraph import CT_P
            from docx.table import Table
            from docx.text.paragraph import Paragraph
            
            segments = []
            for child in doc.element.body:
                if isinstance(child, CT_P):
                    para = Paragraph(child, doc)
                    if para.text.strip():
                        segments.append({"type": "paragraph", "obj": para})
                elif isinstance(child, CT_Tbl):
                    table = Table(child, doc)
                    seen_in_row = {}
                    for row_idx, row in enumerate(table.rows):
                        if row_idx not in seen_in_row:
                            seen_in_row[row_idx] = set()
                        for cell in row.cells:
                            text = cell.text.strip()
                            if text and text not in seen_in_row[row_idx]:
                                seen_in_row[row_idx].add(text)
                                if cell.paragraphs:
                                    segments.append({"type": "cell", "obj": cell.paragraphs[0]})
            
            # Apply mappings
            updated_count = 0
            for path, value in st.session_state.field_values.items():
                if value and path in mapping_config:
                    doc_loc = mapping_config[path]["document_location"]
                    idx = doc_loc["index"]
                    
                    if idx < len(segments):
                        para = segments[idx]["obj"]
                        
                        # Preserve formatting
                        if para.runs:
                            first_run = para.runs[0]
                            bold = first_run.bold
                            italic = first_run.italic
                            font_name = first_run.font.name
                            font_size = first_run.font.size
                            
                            para.clear()
                            run = para.add_run(str(value))
                            run.bold = bold
                            run.italic = italic
                            if font_name:
                                run.font.name = font_name
                            if font_size:
                                run.font.size = font_size
                        else:
                            para.text = str(value)
                        
                        updated_count += 1
            
            st.success(f"✅ Generated document! Updated {updated_count} fields.")
            
            # Save to buffer
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            # Download buttons
            col_a, col_b = st.columns(2)
            
            with col_a:
                st.download_button(
                    "📥 Download Word Document",
                    data=buffer,
                    file_name="generated_document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            with col_b:
                json_str = json.dumps(output_json, indent=2)
                st.download_button(
                    "📥 Download JSON Data",
                    data=json_str,
                    file_name="document_data.json",
                    mime="application/json"
                )
    
    with col2:
        st.info("👈 Click to generate the Word document with your data")

else:
    st.info("👆 Upload a Word template and mapping configuration to begin")
    
    with st.expander("ℹ️ How to use"):
        st.markdown("""
        ### Step-by-step guide:
        
        1. **Upload Word Template**: The `.docx` file you want to populate
        2. **Upload Mapping Config**: The `.json` file from the Interactive Mapper
        3. **Upload Existing Data (Optional)**: Pre-fill fields with a `.json` data file
        
        ### Features:
        - 📅 **Date fields**: Use calendar picker for easy date selection
        - 📧 **Email fields**: Automatic format validation
        - 📁 **Grouped by section**: Fields organized by schema groups
        - 👁️ **Preview JSON**: See the complete JSON structure before generating
        - 💾 **Download both**: Get the Word document AND the JSON data
        
        ### Workflow:
        1. Fill in all required fields (or edit existing values)
        2. Preview the JSON structure
        3. Click "Generate Document"
        4. Download the populated Word document
        5. Optionally download the JSON for future use
        """)


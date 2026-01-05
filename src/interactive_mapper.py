import streamlit as st
import pandas as pd
from docx import Document
import json
import re
from io import BytesIO
import hashlib

from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

st.set_page_config(page_title="Interactive Pilot Brief Mapper", layout="wide", initial_sidebar_state="expanded")

# Add custom CSS for better hover effects
st.markdown("""
<style>
    .selectable-text {
        transition: all 0.2s ease;
    }
    .selectable-text:hover {
        transform: scale(1.02);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .mapped {
        background: #d4edda !important;
        border-color: #28a745 !important;
    }
    .current-target {
        border: 3px dashed #007bff !important;
        animation: pulse 1.5s infinite;
    }
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.7; }
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'mappings' not in st.session_state:
    st.session_state.mappings = {}
if 'current_field' not in st.session_state:
    st.session_state.current_field = None
if 'word_doc' not in st.session_state:
    st.session_state.word_doc = None
if 'json_fields' not in st.session_state:
    st.session_state.json_fields = []

# ============================================================
# Helper Functions
# ============================================================

def resolve_schema_ref(root_schema, ref_path):
    """Resolve a $ref path in JSON schema."""
    if not ref_path.startswith("#/"):
        return None
    parts = ref_path[2:].split("/")
    current = root_schema
    for part in parts:
        if not isinstance(current, dict):
            return None
        current = current.get(part)
        if current is None:
            return None
    return current

def extract_json_fields(schema_data, prefix="", section="", root=None, group=None):
    """Extract all required fields from JSON schema."""
    if root is None:
        root = schema_data
    if not isinstance(schema_data, dict):
        return []
    
    fields = []
    required_set = set(schema_data.get("required", []))
    include_all = len(required_set) == 0
    
    if "properties" in schema_data:
        for k, v in schema_data["properties"].items():
            if not include_all and k not in required_set:
                continue
            
            path = f"{prefix}.{k}" if prefix else k
            current_section = k if k.startswith("section_") else section
            current_group = group or k
            
            # Handle anyOf/allOf/oneOf
            combo = v.get("anyOf") or v.get("allOf") or v.get("oneOf")
            if combo:
                for option in combo:
                    if option.get("type") == "null":
                        continue
                    if "$ref" in option:
                        ref = resolve_schema_ref(root, option["$ref"])
                        if ref:
                            fields += extract_json_fields(ref, path, current_section, root, current_group)
                continue
            
            if "$ref" in v:
                ref = resolve_schema_ref(root, v["$ref"])
                if ref:
                    fields += extract_json_fields(ref, path, current_section, root, current_group)
            elif v.get("type") == "object":
                fields += extract_json_fields(v, path, current_section, root, current_group)
            elif v.get("type") == "array":
                items = v.get("items")
                if isinstance(items, dict) and "$ref" in items:
                    ref = resolve_schema_ref(root, items["$ref"])
                    if ref:
                        fields += extract_json_fields(ref, f"{path}[]", current_section, root, current_group)
            else:
                # Leaf field
                fields.append({
                    "path": path,
                    "field_name": k,
                    "section": current_section,
                    "group": current_group,
                    "type": v.get("type", "unknown")
                })
    
    return fields

def parse_word_document(doc):
    """Parse Word document into a structured format for display."""
    elements = []
    current_section = None
    
    for child in doc.element.body:
        if isinstance(child, CT_P):
            para = Paragraph(child, doc)
            text = para.text.strip()
            if not text:
                continue
            
            style = para.style.name if para.style else ""
            is_heading = style.lower().startswith("heading")
            
            # Check for section markers
            m = re.search(r"section\s*(\d+)", text, re.I)
            if m:
                current_section = f"section_{m.group(1).zfill(2)}"
                is_heading = True
            
            elements.append({
                "type": "paragraph",
                "text": text,
                "is_heading": is_heading,
                "section": current_section,
                "style": style,
                "index": len(elements)
            })
        
        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            headers = [cell.text.strip() for cell in table.rows[0].cells if cell.text.strip()]
            
            table_data = []
            for row_idx, row in enumerate(table.rows):
                row_data = []
                for cell_idx, cell in enumerate(row.cells):
                    row_data.append({
                        "text": cell.text.strip(),
                        "row": row_idx,
                        "col": cell_idx
                    })
                table_data.append(row_data)
            
            elements.append({
                "type": "table",
                "headers": headers,
                "data": table_data,
                "section": current_section,
                "index": len(elements)
            })
    
    return elements

def make_element_id(element_index, row=None, col=None):
    """Create a unique ID for an element."""
    if row is not None and col is not None:
        return f"elem_{element_index}_r{row}_c{col}"
    return f"elem_{element_index}"

def render_document_element(element, element_index):
    """Render a document element with direct click handlers."""
    if element["type"] == "paragraph":
        text = element["text"]
        elem_id = make_element_id(element_index)
        style = "font-size: 20px; font-weight: bold;" if element["is_heading"] else "font-size: 14px;"
        
        # Check if this element is mapped
        mapped_to = None
        for field_path, mapping in st.session_state.mappings.items():
            if mapping.get("element_index") == element_index:
                mapped_to = field_path
                break
        
        classes = "selectable-text"
        if mapped_to:
            classes += " mapped"
        
        bg_color = "#d4edda" if mapped_to else "#ffffff"
        border = "2px solid #28a745" if mapped_to else "1px solid #ddd"
        
        # Create clickable paragraph
        para_html = f"""
        <div id="{elem_id}" 
             class="{classes}"
             data-element="{element_index}"
             style="padding: 15px; margin: 8px 0; background: {bg_color}; 
                    border: {border}; border-radius: 6px; cursor: pointer; {style}"
             onclick="handleClick('paragraph', {element_index}, null, null)">
            {text}
            {f'<div style="color: #28a745; font-size: 11px; margin-top: 5px;">✓ Mapped to: <b>{mapped_to.split(".")[-1]}</b></div>' if mapped_to else ''}
        </div>
        """
        return para_html
    
    elif element["type"] == "table":
        table_html = "<table style='width: 100%; border-collapse: collapse; margin: 15px 0; background: white;'>"
        
        for row_idx, row in enumerate(element["data"]):
            table_html += "<tr>"
            for col_idx, cell in enumerate(row):
                elem_id = make_element_id(element_index, row_idx, col_idx)
                
                # Check if this cell is mapped
                mapped_to = None
                for field_path, mapping in st.session_state.mappings.items():
                    if (mapping.get("element_index") == element_index and 
                        mapping.get("row") == row_idx and 
                        mapping.get("col") == col_idx):
                        mapped_to = field_path
                        break
                
                classes = "selectable-text"
                if mapped_to:
                    classes += " mapped"
                
                bg_color = "#d4edda" if mapped_to else "#ffffff"
                border_style = "2px solid #28a745" if mapped_to else "1px solid #999"
                font_weight = "bold" if row_idx == 0 else "normal"
                
                cell_text = cell['text'] if cell['text'] else "&nbsp;"
                mapped_label = f'<div style="color: #28a745; font-size: 10px; margin-top: 3px;">✓ {mapped_to.split(".")[-1]}</div>' if mapped_to else ''
                
                table_html += f"""
                <td id="{elem_id}"
                    class="{classes}"
                    data-element="{element_index}" data-row="{row_idx}" data-col="{col_idx}"
                    style='padding: 12px; border: {border_style}; background: {bg_color}; 
                           font-weight: {font_weight}; cursor: pointer; min-width: 100px;'
                    onclick="handleClick('table_cell', {element_index}, {row_idx}, {col_idx})">
                    {cell_text}
                    {mapped_label}
                </td>
                """
            table_html += "</tr>"
        
        table_html += "</table>"
        return table_html
    
    return ""

# ============================================================
# Main App
# ============================================================

st.title("🎯 Interactive Pilot Brief Mapper")
st.markdown("Click on text in the document to map it to JSON fields from the sidebar")

# File uploaders in columns
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("📄 Upload Word Template", type="docx")
with col2:
    schema_file = st.file_uploader("📋 Upload JSON Schema", type="json")

if template_file and schema_file:
    # Load files
    doc = Document(template_file)
    schema_data = json.load(schema_file)
    
    # Parse document and extract fields
    if st.session_state.word_doc is None:
        st.session_state.word_doc = parse_word_document(doc)
        st.session_state.json_fields = extract_json_fields(schema_data)
    
    # Sidebar: JSON Fields to map
    with st.sidebar:
        st.header("📋 JSON Fields")
        st.markdown("Click a field, then click text in the document to map it")
        
        # Group by section/group
        fields_by_group = {}
        for field in st.session_state.json_fields:
            group = field.get("group", "ungrouped")
            fields_by_group.setdefault(group, []).append(field)
        
        for group_name, fields in sorted(fields_by_group.items()):
            with st.expander(f"**{group_name}**", expanded=True):
                for field in fields:
                    is_current = st.session_state.current_field == field["path"]
                    is_mapped = field["path"] in st.session_state.mappings
                    
                    button_type = "primary" if is_current else ("secondary" if not is_mapped else "secondary")
                    icon = "🎯" if is_current else ("✅" if is_mapped else "⭕")
                    
                    if st.button(
                        f"{icon} {field['field_name']}", 
                        key=f"field_{field['path']}",
                        use_container_width=True,
                        type=button_type
                    ):
                        st.session_state.current_field = field["path"]
                        st.rerun()
        
        st.markdown("---")
        
        # Show current selection
        if st.session_state.current_field:
            st.info(f"**Selected:** `{st.session_state.current_field}`")
            st.caption("Now click on text in the document to map it")
        
        st.markdown("---")
        
        # Mapping summary
        st.subheader("📊 Mapping Progress")
        total = len(st.session_state.json_fields)
        mapped = len(st.session_state.mappings)
        st.progress(mapped / total if total > 0 else 0)
        st.write(f"{mapped} / {total} fields mapped")
        
        if st.button("🗑️ Clear All Mappings", use_container_width=True):
            st.session_state.mappings = {}
            st.rerun()
    
    # Main content: Rendered document
    st.markdown("### 📄 Document (Click directly on text to map)")
    
    if st.session_state.current_field:
        st.info(f"🎯 **Currently mapping:** `{st.session_state.current_field}` - Click on any text below")
    else:
        st.warning("👈 Select a field from the sidebar first, then click on text in the document")
    
    # Check for click via query params
    query_params = st.query_params
    if "clicked_element" in query_params:
        elem_idx = int(query_params["clicked_element"])
        click_type = query_params.get("click_type", "paragraph")
        
        if st.session_state.current_field:
            if click_type == "paragraph":
                # Find the element
                element = st.session_state.word_doc[elem_idx]
                st.session_state.mappings[st.session_state.current_field] = {
                    "element_index": elem_idx,
                    "type": "paragraph",
                    "text": element["text"]
                }
            elif click_type == "table_cell":
                row_idx = int(query_params["row"])
                col_idx = int(query_params["col"])
                element = st.session_state.word_doc[elem_idx]
                cell_text = element["data"][row_idx][col_idx]["text"]
                st.session_state.mappings[st.session_state.current_field] = {
                    "element_index": elem_idx,
                    "type": "table_cell",
                    "row": row_idx,
                    "col": col_idx,
                    "text": cell_text
                }
            
            # Clear selection and query params
            st.session_state.current_field = None
            st.query_params.clear()
            st.rerun()
    
    # Render document
    doc_html = """
    <div style='background: white; padding: 30px; border: 1px solid #ddd; border-radius: 8px; 
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;'>
    """
    
    for element in st.session_state.word_doc:
        doc_html += render_document_element(element, element["index"])
    
    doc_html += "</div>"
    
    # JavaScript to handle clicks and update URL
    js_code = """
    <script>
    function handleClick(type, elementIdx, row, col) {
        // Build URL with query params
        let url = window.location.href.split('?')[0];
        url += '?clicked_element=' + elementIdx + '&click_type=' + type;
        if (row !== null) url += '&row=' + row;
        if (col !== null) url += '&col=' + col;
        
        // Reload with new params
        window.top.location.href = url;
    }
    
    // Add visual feedback on hover
    document.querySelectorAll('.selectable-text').forEach(el => {
        el.addEventListener('mouseenter', function() {
            if (!this.classList.contains('mapped')) {
                this.style.background = '#e3f2fd';
            }
        });
        el.addEventListener('mouseleave', function() {
            if (!this.classList.contains('mapped')) {
                this.style.background = '#ffffff';
            }
        });
    });
    </script>
    """
    
    st.components.v1.html(doc_html + js_code, height=600, scrolling=True)
    
    # View current mappings
    st.markdown("---")
    st.subheader("📋 Current Mappings")
    
    if st.session_state.mappings:
        mapping_data = []
        for field_path, mapping in st.session_state.mappings.items():
            field_name = field_path.split(".")[-1]
            text_preview = mapping["text"][:50] + "..." if len(mapping["text"]) > 50 else mapping["text"]
            mapping_data.append({
                "JSON Field": field_name,
                "Full Path": field_path,
                "Mapped Text": text_preview,
                "Type": mapping["type"]
            })
        
        df = pd.DataFrame(mapping_data)
        st.dataframe(df, use_container_width=True, hide_index=True)
        
        # Clear individual mappings
        cols = st.columns(5)
        for idx, (field_path, _) in enumerate(list(st.session_state.mappings.items())):
            with cols[idx % 5]:
                if st.button(f"🗑️ {field_path.split('.')[-1]}", key=f"clear_{field_path}"):
                    del st.session_state.mappings[field_path]
                    st.rerun()
    else:
        st.info("No mappings yet. Select a field from the sidebar and click on text in the document.")
    
    # Extract button
    st.markdown("---")
    if st.button("🚀 Extract Data to JSON", type="primary", disabled=len(st.session_state.mappings) == 0):
        # Build JSON from mappings
        json_output = {}
        
        for field_path, mapping in st.session_state.mappings.items():
            # Set nested value
            parts = field_path.replace("[]", "").split(".")
            current = json_output
            for part in parts[:-1]:
                if part not in current:
                    current[part] = {}
                current = current[part]
            current[parts[-1]] = mapping["text"]
        
        st.success("✅ Data extracted successfully!")
        st.json(json_output)
        
        # Download button
        json_str = json.dumps(json_output, indent=2)
        st.download_button(
            "📥 Download JSON",
            data=json_str,
            file_name="extracted_data.json",
            mime="application/json"
        )

else:
    st.info("👆 Upload both a Word template and JSON schema to begin")


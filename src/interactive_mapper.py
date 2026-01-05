import streamlit as st
import pandas as pd
from docx import Document
import json
import re

from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

st.set_page_config(page_title="JSON to Document Mapper", layout="wide")

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
    """Parse Word document into text segments."""
    segments = []
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
            
            segments.append({
                "text": text,
                "type": "heading" if is_heading else "paragraph",
                "section": current_section,
                "style": style
            })
        
        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            
            # Track unique text in each row to avoid merged cell duplicates
            # but keep the same text if it appears in different rows
            seen_in_row = {}
            
            for row_idx, row in enumerate(table.rows):
                # Reset for each row
                if row_idx not in seen_in_row:
                    seen_in_row[row_idx] = set()
                
                for cell_idx, cell in enumerate(row.cells):
                    text = cell.text.strip()
                    if text:
                        # Create a unique key for this text + row
                        cell_key = (row_idx, text)
                        
                        # Skip if we've seen this exact text in this row (merged cell)
                        if text in seen_in_row[row_idx]:
                            continue
                        
                        seen_in_row[row_idx].add(text)
                        
                        is_header = row_idx == 0
                        segments.append({
                            "text": text,
                            "type": "table_header" if is_header else "table_cell",
                            "section": current_section,
                            "style": f"Table[{row_idx},{cell_idx}]"
                        })
    
    return segments

# ============================================================
# Main App
# ============================================================

st.title("📋 JSON to Document Mapper")
st.markdown("Map JSON schema fields to Word document locations - define which text each JSON field should replace.")

# File uploaders
col1, col2 = st.columns(2)
with col1:
    doc_file = st.file_uploader("📄 Upload Word Document", type="docx")
with col2:
    schema_file = st.file_uploader("📋 Upload JSON Schema", type="json")

if doc_file and schema_file:
    # Load schema
    schema_data = json.load(schema_file)
    json_fields = extract_json_fields(schema_data)
    
    # Group JSON fields
    fields_by_group = {}
    for field in json_fields:
        group = field.get("group", "other")
        fields_by_group.setdefault(group, []).append(field)
    
    # Parse document
    doc = Document(doc_file)
    segments = parse_word_document(doc)
    
    st.success(f"✓ Found {len(segments)} text blocks and {len(json_fields)} JSON fields")
    
    # Create text block options for dropdowns with better formatting
    text_options = ["(Not Mapped)"]
    text_lookup = {}
    
    for idx, seg in enumerate(segments):
        # Simpler format: just ID and text (no icons)
        text_preview = seg["text"][:60] + "..." if len(seg["text"]) > 60 else seg["text"]
        label = f"{idx:03d} | {text_preview}"
        text_options.append(label)
        text_lookup[label] = {
            "index": idx,
            "text": seg["text"],
            "type": seg["type"],
            "section": seg["section"]
        }
    
    # Create DataFrame with JSON fields to map
    df_data = []
    for field in json_fields:
        df_data.append({
            "JSON Field": field["field_name"],
            "Full Path": field["path"],
            "Group": field["group"],
            "Type": field["type"],
            "Mapped To": "(Not Mapped)"
        })
    
    df = pd.DataFrame(df_data)
    
    # Function to render document preview
    def render_document_preview_with_mappings(segments, segment_to_field):
        """Render the Word document with highlighting for mapped sections."""
        html = """
        <div style='background: white; padding: 30px; border: 2px solid #ddd; 
                    border-radius: 8px; height: 600px; overflow-y: scroll;
                    font-family: "Calibri", "Arial", sans-serif;'>
        """
        
        for idx, seg in enumerate(segments):
            text = seg["text"]
            
            # Check if this segment has a JSON field mapped to it
            is_mapped = idx in segment_to_field
            json_field = segment_to_field.get(idx, "")
            
            # Uniform styling - only difference is green for mapped
            if is_mapped:
                bg_color = "#d4edda"
                border_color = "#28a745"
            else:
                bg_color = "#ffffff"
                border_color = "#e9ecef"
            
            # Add type indicator and ID
            type_emoji = {"heading": "📌", "table_header": "📊", "table_cell": "📋", "paragraph": "📝"}
            emoji = type_emoji.get(seg["type"], "📝")
            
            # Add mapped indicator
            mapped_label = ""
            if is_mapped:
                mapped_label = f'<div style="color: #155724; font-size: 11px; margin-top: 4px;">✓ Will be replaced by: <b>{json_field}</b></div>'
            
            html += f"""
            <div style='padding: 12px; margin: 6px 0; background: {bg_color}; 
                       border-left: 4px solid {border_color}; border-radius: 4px;
                       font-size: 14px; font-weight: normal;'>
                <span style='color: #6c757d; font-size: 11px;'>[{idx:03d}] {emoji}</span> {text}
                {mapped_label}
            </div>
            """
        
        html += "</div>"
        return html
    
    # Search/filter
    st.markdown("---")
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        search = st.text_input("🔍 Search JSON field names", "")
    with col2:
        group_filter = st.multiselect("Filter by group", 
                                     sorted(fields_by_group.keys()),
                                     default=sorted(fields_by_group.keys()))
    with col3:
        show_only_unmapped = st.checkbox("Show only unmapped", value=False)
    
    # Apply filters
    df_filtered = df.copy()
    if search:
        df_filtered = df_filtered[df_filtered["JSON Field"].str.contains(search, case=False, na=False)]
    if group_filter:
        df_filtered = df_filtered[df_filtered["Group"].isin(group_filter)]
    if show_only_unmapped:
        df_filtered = df_filtered[df_filtered["Mapped To"] == "(Not Mapped)"]
    
    st.markdown("---")
    
    # Two column layout: Mapping table on left, Document preview on right
    col_table, col_doc = st.columns([1, 1])
    
    with col_table:
        st.markdown("### 📝 Map JSON Fields to Document")
        st.caption("For each JSON field, select where in the document it should go")
        
        # Interactive editor
        edited_df = st.data_editor(
            df_filtered,
            column_config={
                "JSON Field": st.column_config.TextColumn("Field", width="medium", disabled=True),
                "Full Path": st.column_config.TextColumn("Path", width="medium", disabled=True),
                "Group": st.column_config.TextColumn("Group", width="small", disabled=True),
                "Type": st.column_config.TextColumn("Type", width="small", disabled=True),
                "Mapped To": st.column_config.SelectboxColumn(
                    "Document Location",
                    options=text_options,
                    width="large",
                    help="Select which text in the document this JSON field should replace"
                )
            },
            use_container_width=True,
            hide_index=True,
            height=500,
            key="editor"
        )
        
        # Update original df with edits
        for idx in edited_df.index:
            df.at[idx, "Mapped To"] = edited_df.at[idx, "Mapped To"]
        
        # Refresh preview button
        if st.button("🔄 Refresh Preview"):
            st.rerun()
    
    with col_doc:
        st.markdown("### 📄 Document Preview")
        st.caption("Green = will be replaced by JSON data")
        
        # Build a reverse mapping for preview: segment_idx -> json_field
        segment_to_field = {}
        for _, row in df.iterrows():
            if row["Mapped To"] != "(Not Mapped)":
                location_info = text_lookup.get(row["Mapped To"])
                if location_info:
                    segment_to_field[location_info["index"]] = row["JSON Field"]
        
        # Render document preview with mappings
        doc_html = render_document_preview_with_mappings(segments, segment_to_field)
        st.components.v1.html(doc_html, height=620, scrolling=True)
        
        # Show all text in a list format
        with st.expander("📋 View All Text (List Format)"):
            for idx, seg in enumerate(segments):
                type_emoji = {"heading": "📌", "table_header": "📊", "table_cell": "📋", "paragraph": "📝"}
                emoji = type_emoji.get(seg["type"], "📝")
                
                is_mapped = idx in segment_to_field
                if is_mapped:
                    st.success(f"**[{idx:03d}]** {emoji} {seg['text']} → `{segment_to_field[idx]}`")
                else:
                    st.text(f"[{idx:03d}] {emoji} {seg['text']}")
    
    # Show progress
    mapped_count = (df["Mapped To"] != "(Not Mapped)").sum()
    total_fields = len(df)
    st.progress(mapped_count / total_fields if total_fields > 0 else 0)
    st.caption(f"**Progress:** {mapped_count} / {total_fields} JSON fields mapped")
    
    # Show mapping summary
    with st.expander("📊 Mapping Summary"):
        mapped_fields = df[df["Mapped To"] != "(Not Mapped)"][["JSON Field", "Group", "Mapped To"]]
        if not mapped_fields.empty:
            st.dataframe(mapped_fields, use_container_width=True, hide_index=True)
        else:
            st.info("No mappings yet - select document locations for your JSON fields above")
    
    # Save mapping configuration
    st.markdown("---")
    col1, col2 = st.columns([1, 3])
    
    with col1:
        if st.button("💾 Save Mapping Config", type="primary", disabled=mapped_count == 0):
            # Build mapping configuration
            mapping_config = {}
            
            for _, row in df.iterrows():
                if row["Mapped To"] != "(Not Mapped)":
                    json_path = row["Full Path"]
                    location_info = text_lookup.get(row["Mapped To"])
                    
                    if location_info:
                        mapping_config[json_path] = {
                            "field_name": row["JSON Field"],
                            "group": row["Group"],
                            "document_location": {
                                "index": location_info["index"],
                                "text": location_info["text"],
                                "type": location_info["type"],
                                "section": location_info["section"]
                            }
                        }
            
            st.success(f"✅ Mapping configuration created! {len(mapping_config)} fields mapped.")
            st.json(mapping_config)
            
            # Download button
            config_str = json.dumps(mapping_config, indent=2)
            st.download_button(
                "📥 Download Mapping Config",
                data=config_str,
                file_name="mapping_config.json",
                mime="application/json"
            )
            
            st.info("💡 Use this mapping config to populate Word documents with JSON data later!")
    
    with col2:
        if mapped_count == 0:
            st.info("👈 Map some JSON fields to document locations first, then save the configuration")

else:
    st.info("👆 Upload both a Word document and JSON schema to begin mapping")
    
    # Show example
    with st.expander("ℹ️ How it works"):
        st.markdown("""
        1. **Upload files**: Word document (template) + JSON schema (defines fields)
        2. **Map JSON fields**: For each required JSON field, select which text in the document it should replace
        3. **Save mapping**: Download the mapping configuration file
        4. **Use later**: Load this config to populate Word templates with actual JSON data
        
        **What you're creating:**
        - A mapping that says "put `title` field here, put `date_written` there", etc.
        - This config can be reused for any document using the same template
        
        **Tips:**
        - Look at the document preview (right side) to see what text exists
        - The dropdown shows `[ID] icon text_preview` for easy identification
        - Green text in preview = already mapped to a JSON field
        - Use filters to focus on specific groups (section_01, section_02, etc.)
        """)

import streamlit as st
import pandas as pd
from docx import Document
import json
import re

from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from mapper import normalize_header, map_row_data

st.set_page_config(page_title="Flexible Exports Pilot", layout="wide", initial_sidebar_state="collapsed")


def walk_container(container, path):
    nodes = []

    # Paragraphs in this container
    for i, p in enumerate(container.paragraphs):
        if p.text.strip():
            nodes.append({
                "type": "paragraph",
                "text": p.text,
                "style": p.style.name if p.style else None,
                "path": f"{path}/p[{i}]",
                "obj": p
            })

    # Tables in this container
    for ti, table in enumerate(container.tables):
        table_path = f"{path}/table[{ti}]"
        table_node = {
            "type": "table",
            "path": table_path,
            "rows": []
        }

        for ri, row in enumerate(table.rows):
            row_node = {"type": "row", "cells": []}

            for ci, cell in enumerate(row.cells):
                cell_path = f"{table_path}/row[{ri}]/cell[{ci}]"
                cell_node = {
                    "type": "cell",
                    "path": cell_path,
                    "children": walk_container(cell, cell_path)
                }
                row_node["cells"].append(cell_node)

            table_node["rows"].append(row_node)

        nodes.append(table_node)

    return nodes

def collect_text_nodes(doc_structure):
    """
    Collect only text-bearing items (headings, paragraphs, table headers) with section info.
    Skips blanks and returns label + info for selection.
    """
    items = []
    for idx, item in enumerate(doc_structure):
        section = item.get("section")
        if item["type"] in ("heading", "paragraph"):
            text = item["text"].strip()
            if not text:
                continue
            label = f"{text[:60]}{'...' if len(text) > 60 else ''}"
            items.append({
                "label": label,
                "section": section,
                "text": text,
                "path": f"item[{idx}]"
            })
        elif item["type"] == "table" and item.get("headers"):
            for h_idx, header in enumerate(item["headers"]):
                label = f"{header} (Table header)"
                items.append({
                    "label": label,
                    "section": section,
                    "text": header,
                    "path": f"table[{idx}]/header[{h_idx}]"
                })
    return items

def render_node(node):
    """
    Render a node from the walk_container DOM tree into Streamlit.
    Shows paragraphs and tables recursively with paths for context.
    """
    if node["type"] == "paragraph":
        st.markdown(f"**{node['path']}**")
        st.text(node["text"])

    elif node["type"] == "table":
        with st.expander(f"📊 {node['path']}"):
            for row in node["rows"]:
                for cell in row["cells"]:
                    with st.expander(cell["path"]):
                        for child in cell["children"]:
                            render_node(child)

# ============================================================
# Word document model (THIS IS THE FIX)
# ============================================================

def iter_block_items(doc):
    for child in doc.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def extract_document_structure(doc):
    structure = []
    current_section = None

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue

            style = block.style.name if block.style else ""

            is_heading = False
            level = None

            if style.lower().startswith("heading"):
                is_heading = True
                m = re.search(r"(\d+)", style)
                if m:
                    level = int(m.group(1))

            m2 = re.search(r"section\s*(\d+)", text, re.I)
            if m2:
                is_heading = True
                current_section = f"section_{m2.group(1).zfill(2)}"

            structure.append({
                "type": "heading" if is_heading else "paragraph",
                "text": text,
                "heading_level": level,
                "style": style,
                "section": current_section
            })

        elif isinstance(block, Table):
            headers = [cell.text.strip() for cell in block.rows[0].cells if cell.text.strip()]
            structure.append({
                "type": "table",
                "table": block,
                "rows": len(block.rows),
                "cols": len(block.columns),
                "headers": headers,
                "section": current_section
            })

    return structure

# ============================================================
# JSON Schema logic (unchanged)
# ============================================================

def resolve_schema_ref(root_schema, ref_path):
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

def extract_json_paths(schema_data, prefix="", section="", root=None):
    """
    Extract all leaf fields from the JSON schema with proper section context.
    Includes strings like title, date_written, date_due across all sections.
    """
    if root is None:
        root = schema_data
    if not isinstance(schema_data, dict):
        return []

    paths = []
    if "properties" in schema_data:
        for k, v in schema_data["properties"].items():
            path = f"{prefix}.{k}" if prefix else k
            current_section = k if k.startswith("section_") else section

            # Handle anyOf/allOf/oneOf with $ref inside
            combo = v.get("anyOf") or v.get("allOf") or v.get("oneOf")
            if combo:
                for option in combo:
                    if "$ref" in option:
                        ref = resolve_schema_ref(root, option["$ref"])
                        if ref:
                            paths += extract_json_paths(ref, path, current_section, root)
                continue

            if "$ref" in v:
                ref = resolve_schema_ref(root, v["$ref"])
                if ref:
                    paths += extract_json_paths(ref, path, current_section, root)
            elif v.get("type") == "object":
                paths += extract_json_paths(v, path, current_section, root)
            elif v.get("type") == "array":
                items = v.get("items")
                if isinstance(items, dict) and "$ref" in items:
                    ref = resolve_schema_ref(root, items["$ref"])
                    if ref:
                        paths += extract_json_paths(ref, path, current_section, root)
                elif isinstance(items, dict):
                    paths += extract_json_paths(items, path, current_section, root)
            else:
                # Leaf node (only include strings for text mapping)
                if v.get("type") == "string":
                    paths.append({"path": path, "field_name": k, "section": current_section})

    return paths

# ============================================================
# Streamlit App
# ============================================================

st.title("Flexible Exports Pilot")

template_file = st.file_uploader("Upload Word Template", type="docx")
schema_file = st.file_uploader("Upload JSON Schema", type="json")

if template_file and schema_file:
    schema_data = json.load(schema_file)
    doc = Document(template_file)

    doc_structure = extract_document_structure(doc)
    doc_tree = walk_container(doc, "doc")
    tables = [x for x in doc_structure if x["type"] == "table"]

    # Map table path -> section for DOM tagging
    table_section_map = {f"doc/table[{i}]": t["section"] for i, t in enumerate(tables)}

    st.success(f"✓ Found {len(tables)} table(s)")

    with st.expander("📄 View Document Structure"):
        for item in doc_structure:
            if item["type"] == "heading":
                st.markdown(f"### 📌 {item['text']}")
            elif item["type"] == "paragraph":
                st.text(item["text"])
            else:
                st.markdown(f"📊 Table ({item['rows']}×{item['cols']}) — Section: {item['section']}")
                if item["headers"]:
                    st.caption(", ".join(item["headers"]))

    # Full DOM walk (paragraphs, tables, cells)
    with st.expander("📄 Full Word DOM", expanded=False):
        st.caption("Nested view of all paragraphs, tables, rows, and cells with paths")
        for n in doc_tree:
            render_node(n)

    # Extract Word columns
    word_columns = []
    for idx, t in enumerate(tables):
        for cell in t["table"].rows[0].cells:
            if cell.text.strip():
                word_columns.append({
                    "column_name": cell.text.strip(),
                    "table_index": idx,
                    "section": t["section"]
                })

    json_paths = extract_json_paths(schema_data)

    word_col_display = []
    word_lookup = {}
    for c in word_columns:
        label = f"{c['column_name']} ({c['section'] or 'No section'})"
        word_col_display.append(label)
        word_lookup[label] = {
            **c,
            "type": "header_cell",
            "path": f"table[{c['table_index']}]/header/{c['column_name']}"
        }

    # Text options (paragraphs, headings, table headers) by section
    text_nodes = collect_text_nodes(doc_structure)
    text_by_section = {}
    for n in text_nodes:
        text_by_section.setdefault(n["section"], []).append(n)

    st.subheader("Review Mappings (JSON on left, Word text on right)")

    selections = {}

    # Group fields by section for table-style UI
    fields_by_section = {}
    for j in json_paths:
        fields_by_section.setdefault(j["section"], []).append(j)

    for section, fields in fields_by_section.items():
        st.markdown(f"#### Section: {section or 'N/A'}")
        options = text_by_section.get(section, [])
        option_labels = [None] + [o["label"] for o in options]
        option_lookup = {o["label"]: o for o in options}

        data = []
        for j in fields:
            # auto-guess exact text match
            guess_label = None
            for o in options:
                if normalize_header(o["text"]) == normalize_header(j["field_name"]):
                    guess_label = o["label"]
                    break
            data.append({"JSON Field": j["path"], "Word Text": guess_label})

        df = pd.DataFrame(data)

        edited = st.data_editor(
            df,
            column_config={
                "Word Text": st.column_config.SelectboxColumn(
                    options=option_labels,
                    help="Select text from the same section of the Word document"
                ),
                "JSON Field": st.column_config.TextColumn(disabled=True)
            },
            hide_index=True
        )

        # Collect selections
        for _, row in edited.iterrows():
            if pd.notna(row["Word Text"]):
                selections[row["JSON Field"]] = option_lookup.get(row["Word Text"])

    if st.button("Generate"):
        final = selections
        st.json(final)

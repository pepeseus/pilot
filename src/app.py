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

def resolve_schema_ref(schema_data, ref_path):
    if not ref_path.startswith("#/"):
        return None
    parts = ref_path[2:].split("/")
    current = schema_data
    for part in parts:
        current = current.get(part)
        if current is None:
            return None
    return current

def extract_json_paths(schema_data, prefix="", section=""):
    paths = []
    if "properties" in schema_data:
        for k, v in schema_data["properties"].items():
            path = f"{prefix}.{k}" if prefix else k
            new_section = k if k.startswith("section_") else section

            if "$ref" in v:
                ref = resolve_schema_ref(schema_data, v["$ref"])
                paths += extract_json_paths(ref, path, new_section)
            elif v.get("type") == "object":
                paths += extract_json_paths(v, path, new_section)
            elif v.get("type") == "array":
                items = v["items"]
                if "$ref" in items:
                    ref = resolve_schema_ref(schema_data, items["$ref"])
                    paths += extract_json_paths(ref, path, new_section)
            else:
                paths.append({"path": path, "field_name": k, "section": section})

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
        word_lookup[label] = c

    mappings = []
    for j in json_paths:
        guess = None
        for label, c in word_lookup.items():
            if c["section"] == j["section"] and normalize_header(c["column_name"]) == normalize_header(j["field_name"]):
                guess = label
        mappings.append({"JSON Field": j["path"], "Word Column": guess, "Section": j["section"]})

    df = pd.DataFrame(mappings)

    edited = st.data_editor(df, column_config={
        "Word Column": st.column_config.SelectboxColumn(options=[None] + word_col_display)
    }, hide_index=True)

    if st.button("Generate"):
        final = {}
        for _, row in edited.iterrows():
            if pd.notna(row["Word Column"]):
                final[row["JSON Field"]] = word_lookup[row["Word Column"]]
        st.json(final)

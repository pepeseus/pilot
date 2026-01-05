# PilotBrief

A system to populate Word document templates with JSON data, designed to support frequent template updates with minimal code changes.

## Project Structure

* `data/`: Static assets (Templates, Schemas, Example JSON).
* `src/`: Application source code.
    * `schema.py`: Pydantic schema definitions for data validation.
    * `interactive_mapper.py`: Streamlit app for mapping JSON fields to Word document sections.
    * `document_generator.py`: Streamlit app for generating populated Word documents from JSON data.

## Setup Instructions

1.  **Create a Virtual Environment**
    ```bash
    python -m venv .venv
    .venv\Scripts\activate  # On Windows
    # source .venv/bin/activate  # On macOS/Linux
    ```

2.  **Install Dependencies**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Interactive Mapper** - Map JSON fields to document sections
    ```bash
    streamlit run src/interactive_mapper.py
    ```
    This tool helps you create and edit mappings between JSON schema fields and Word document sections. It automatically detects headings in your document and allows you to assign JSON fields to specific sections.

2.  **Document Generator** - Populate templates with JSON data
    ```bash
    streamlit run src/document_generator.py
    ```
    This tool takes a Word template, a mapping file, and JSON data to generate a populated document. It handles various data types including dates, emails, and nested objects.

3.  **Run Tests** (if available)
    ```bash
    pytest tests/
    ```

## Key Features
* **Flexible Section Naming:** Headings in your Word document can be named anything - the system automatically detects them.
* **Format Recognition:** Automatically detects and displays dates (📅) and email addresses (📧) based on JSON schema format annotations.
* **Dynamic Mapping:** Create and edit field-to-section mappings through an interactive web interface.
* **Template Support:** Works with any Word document template structure.
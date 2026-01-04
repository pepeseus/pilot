# pilot

A system to populate Word document templates with JSON data, designed to support frequent template updates with minimal code changes.

## Project Structure

* `data/`: Static assets (Templates, Schemas, Example JSON).
* `src/`: Application source code.
    * `schema.py`: Pydantic definitions.
    * `mapper.py`: Logic to map JSON data to Word table columns.
    * `doc_generator.py`: Main logic to write the .docx file.
* `output/`: Generated documents (ignored by Git).
* `tests/`: Unit tests.

## Setup Instructions

1.  **Create a Virtual Environment**
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate  # On Windows: .venv\Scripts\activate
    ```

2.  **Install Dependencies**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Run the Generator** (Once implemented)
    ```bash
    python src/main.py
    ```

2.  **Run Tests**
    ```bash
    pytest tests/
    ```

## Key Assumptions
* **Header Mapping:** The system assumes Word table headers (e.g., "Phone No.") broadly match JSON keys (e.g., `phone_no`).
* **Polymorphism:** Section 03 handles multiple step types (`Standard`, `DateTime`, `Subtitle`) requiring dynamic row handling.
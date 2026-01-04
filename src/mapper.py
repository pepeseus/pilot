import re

def normalize_header(header_text):
    """
    Converts a Word table header into a snake_case JSON key.
    Example: "Phone No." -> "phone_no"
             "Active Employee?" -> "active_employee"
    """
    if not header_text:
        return ""
    
    # Lowercase the text
    text = header_text.lower()
    
    # Replace common abbreviations (Optional, but helpful for "Phone No." -> "phone_number")
    # text = text.replace("no.", "number") 
    
    # Remove any characters that aren't letters, numbers, or spaces
    text = re.sub(r'[^a-z0-9\s]', '', text)
    
    # Replace spaces with underscores
    text = re.sub(r'\s+', '_', text)
    
    return text.strip('_')

def map_row_data(headers, json_item):
    """
    Maps a single JSON object to a list of cell values based on headers.
    
    Args:
        headers (list): List of strings from the Word table header row.
                        e.g., ['Name', 'Phone No.', 'Active Employee?']
        json_item (dict): The data object for this row.
                          e.g., {'name': 'Alex', 'phone_no': '123...', ...}
    
    Returns:
        dict: A mapping of column index to value to insert.
    """
    row_mapping = {}
    
    for index, header in enumerate(headers):
        key = normalize_header(header)
        
        # Check if the normalized header exists in the JSON data
        if key in json_item:
            row_mapping[index] = json_item[key]
        else:
            # OPTIONAL: Handle mismatches (e.g., log a warning)
            # print(f"Warning: Could not find data for column '{header}' (key: '{key}')")
            row_mapping[index] = "" 
            
    return row_mapping

# --- Example Usage based on your files ---

# 1. Simulate the headers found in "Section 02" of ExampleDocument.docx
word_table_headers = ["Name", "Phone No.", "Email Address", "Active Employee?"]

# 2. Simulate the JSON data (Note: Keys match the normalized headers)
participant_data = {
    "name": "Alex Stevens",
    "phone_no": "08672 434722",
    "email_address": "astevens@entangl.ai",
    "active_employee": True
}

# 3. Run the mapping
mapped_cells = map_row_data(word_table_headers, participant_data)

print(f"Normalized Keys: {[normalize_header(h) for h in word_table_headers]}")
print("Mapped Data for Word Table:")
for col_idx, value in mapped_cells.items():
    print(f"  Col {col_idx} ({word_table_headers[col_idx]}): {value}")
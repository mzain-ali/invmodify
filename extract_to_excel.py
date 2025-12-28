import fitz
import re
import os
import pandas as pd

# Configuration
INPUT_DIR = "input"
OUTPUT_FILE = "output/all_invoices_data_v6.xlsx"

# Header Keywords
HEADER_UNIT_PRICE = "Unit Price"
HEADER_TOTAL = "Total"
HEADER_QUANTITY = "Quantity"
HEADER_WEIGHT = "Weight" # Search for "Weight" or "Weight(kg)"
HEADER_REF = "TVH"     # Search for "TVH-ref" or "TVH"

# Regex
CURRENCY_REGEX = r"\b[\d\.]+,\d{2}\b"

def parse_euro_decimal(text):
    """Converts '1.234,56' or 'USD 1.234,56' to float"""
    clean = text.replace("USD", "").strip().replace(".", "").replace(",", ".")
    try:
        return float(clean)
    except ValueError:
        return 0.0

def is_same_line(y1, y2, tolerance=3):
    return abs(y1 - y2) < tolerance

def clean_weight_value(text):
    """Parses weight/unit text: '1,234 kg' -> 1.234, '1 set' -> 1.0"""
    if not text: return ""
    text_lower = text.lower()
    
    # Remove known units
    for unit in ["kg", "pieces", "piece", "sets", "set", "pairs", "pair"]:
        text_lower = text_lower.replace(unit, "")
    
    clean = text_lower.strip()
    
    # Replace decimal separator
    clean = clean.replace(",", ".")
    
    # Attempt to extract just the number if there's still junk
    # e.g. "approx 1.2" -> 1.2
    import re
    match = re.search(r"[\d\.]+", clean)
    if match:
        clean = match.group(0)

    try:
        return float(clean)
    except ValueError:
        if not clean: return ""
        return clean

def extract_data_from_pdf(pdf_path, filename):
    doc = fitz.open(pdf_path)
    extracted_rows = []

    for page_num, page in enumerate(doc):
        # 1. Analyze structure & Find Headers
        blocks = page.get_text("dict")["blocks"]
        all_spans = []
        for b in blocks:
            if "lines" not in b: continue
            for l in b["lines"]:
                for s in l["spans"]:
                    all_spans.append(s)

        # Default Coordinates
        x_quantity = 50
        x_weight = 0
        x_ref = 0
        x_unit_price = 450
        x_total = 550
        
        # Scan for actual headers
        for span in all_spans:
            text = span["text"].strip()
            
            if HEADER_QUANTITY in text:
                x_quantity = span["bbox"][0]
            elif HEADER_UNIT_PRICE in text:
                x_unit_price = span["bbox"][0]
            elif HEADER_TOTAL in text and span["bbox"][0] > x_unit_price:
                x_total = span["bbox"][0]
            elif HEADER_WEIGHT in text or "kg" in text.lower() and span["bbox"][1] < 300: # Header usually top half
                # Only trust if it looks like a header (y position check or simple text)
                if x_weight == 0 and span["bbox"][0] > x_quantity and span["bbox"][0] < x_unit_price:
                     x_weight = span["bbox"][0]
            elif HEADER_REF in text:
                if x_ref == 0 and span["bbox"][0] > x_quantity and span["bbox"][0] < x_unit_price:
                    x_ref = span["bbox"][0]

        # Validation / Fallbacks
        # If headers not found, try to guess or use defaults (but careful not to break)
        if x_weight == 0: x_weight = x_quantity + 250 # Fallback guessing
        if x_ref == 0: x_ref = x_weight + 60         # Fallback guessing

        # Calculate Column Boundaries
        # Layout: Qty | Description | Weight | TVH-Ref | Unit Price | Total
        
        col_desc_start = x_quantity + 40 # Give some space for Qty column
        col_weight_start = x_weight - 10 # Padded left of Weight header
        col_ref_start = x_ref - 10       # Padded left of Ref header
        col_price_start = x_unit_price - 20
        col_total_start = x_total - 20

        # Refine boundaries (ensure order is logical: Desc < Weight < Ref < Price)
        if col_desc_start >= col_weight_start: col_weight_start = col_desc_start + 100
        if col_weight_start >= col_ref_start: col_ref_start = col_weight_start + 50
        
        print(f"DEBUG Page {page_num+1}: Columns -> Desc[{col_desc_start:.0f}:{col_weight_start:.0f}] Weight[{col_weight_start:.0f}:{col_ref_start:.0f}] Ref[{col_ref_start:.0f}:{col_price_start:.0f}]")

        # 2. Group by Line Y
        lines = {}
        for span in all_spans:
            y = span["origin"][1]
            found_line = False
            for existing_y in lines.keys():
                if is_same_line(y, existing_y):
                    lines[existing_y].append(span)
                    found_line = True
                    break
            if not found_line:
                lines[y] = [span]

        sorted_ys = sorted(lines.keys())
        
        # 3. Stateful Parsing
        items_buffer = []  # List of dicts
        current_item = None 

        for y in sorted_ys:
            row_spans = lines[y]
            row_spans.sort(key=lambda s: s["bbox"][0])

            # Buckets
            qty_text = ""      # Far Left
            desc_text_accum = "" # Middle Left
            weight_text = ""   # Weight Col
            ref_text = ""      # Ref Col
            price_text = ""    # Price Col
            total_text = ""    # Total Col
            
            for span in row_spans:
                text = span["text"].strip()
                if not text: continue
                x = span["bbox"][0]
                
                # Spatial Bucketing
                if x < col_desc_start:
                    qty_text += text + " "
                elif x >= col_desc_start and x < col_weight_start:
                    desc_text_accum += text + " "
                elif x >= col_weight_start and x < col_ref_start:
                    weight_text += text + " "
                elif x >= col_ref_start and x < col_price_start:
                    ref_text += text + " "
                elif x >= col_price_start and x < col_total_start:
                    price_text += text + " "
                elif x >= col_total_start:
                    total_text += text + " "

            qty_text = qty_text.strip()
            desc_text_accum = desc_text_accum.strip()
            weight_text = weight_text.strip()
            ref_text = ref_text.strip()
            price_text = price_text.strip()
            total_text = total_text.strip()

            # --- HEADER SKIP ---
            if "Unit Price" in price_text or "Weight" in weight_text:
                continue

            # --- IS THIS A MAIN ITEM ROW (Has Price + Total)? ---
            if re.search(CURRENCY_REGEX, price_text) and re.search(CURRENCY_REGEX, total_text):
                if current_item:
                    items_buffer.append(current_item)
                
                # Start NEW Item
                clean_price = price_text
                clean_total = total_text
                
                unit_price = 0.0
                line_total = 0.0
                qty = 0.0
                
                line_total = parse_euro_decimal(clean_total)
                match = re.search(r"^(\d+)\s+([\d\.]+,\d{2})$", clean_price)
                if match:
                    qty = float(match.group(1))
                    unit_price = parse_euro_decimal(match.group(2))
                else:
                    unit_price = parse_euro_decimal(clean_price)
                    if unit_price > 0:
                        qty = round(line_total / unit_price)
                    else:
                        qty = 0.0

                current_item = {
                    "File": filename,
                    "Page": page_num + 1,
                    "Part No": qty_text, 
                    "Harmonized Code": "",
                    "Country of Origin": "",
                    "Description": desc_text_accum, 
                    "Weight": clean_weight_value(weight_text), # CLEANED HERE
                    "TVH Ref": ref_text,   # Explicit Column
                    "Quantity": qty,
                    "Unit Price": unit_price,
                    "Total": line_total
                }
                
                # Logic: If PartNo is TVH Ref, copy it
                if current_item["Part No"].startswith("TVH/"):
                    current_item["TVH Ref"] = current_item["Part No"]
            
            # --- IS THIS A DETAIL ROW? ---
            elif current_item:
                # 1. Capture content from explicit columns
                if weight_text:
                    w_val = clean_weight_value(weight_text)
                    if w_val != "":
                        current_item["Weight"] = w_val
                
                if ref_text:
                    current_item["TVH Ref"] = ref_text
                
                # 2. Metadata in Qty/PartNo column (Harmonized, Country)
                if re.match(r"^\d{8,}$", qty_text):
                    current_item["Harmonized Code"] = qty_text
                
                country_match = re.search(r"([A-Z\s]+)\s+-\s+\d+\s+piece", qty_text + " " + desc_text_accum, re.IGNORECASE)
                if country_match:
                    current_item["Country of Origin"] = country_match.group(1).strip()
                
                # 3. Description Accumulation
                # Add desc_text only if it's not strictly metadata
                # Check if this line looks like just Country or Weight
                is_metadata_line = False
                if re.search(r"^\d{8,}$", qty_text): is_metadata_line = True
                if country_match: is_metadata_line = True
                if "Warranty:" in desc_text_accum: is_metadata_line = True
                
                if not is_metadata_line and desc_text_accum:
                     current_item["Description"] += " " + desc_text_accum

        # Append last item
        if current_item:
            items_buffer.append(current_item)
            
        extracted_rows.extend(items_buffer)

    doc.close()
    return extracted_rows

def main():
    if not os.path.exists(INPUT_DIR):
        print(f"Error: Directory '{INPUT_DIR}' not found.")
        return

    if not os.path.exists("output"):
        os.makedirs("output")

    all_data = []

    files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".pdf")]
    
    if not files:
        print("No PDF files found in input directory.")
        return

    print(f"Found {len(files)} PDF files to process.")

    for filename in files:
        filepath = os.path.join(INPUT_DIR, filename)
        print(f"Processing {filename}...")
        try:
            file_data = extract_data_from_pdf(filepath, filename)
            all_data.extend(file_data)
            print(f"  Extracted {len(file_data)} items.")
        except Exception as e:
            print(f"  Error processing {filename}: {e}")

    if all_data:
        df = pd.DataFrame(all_data)
        
        # Output Order
        cols = [
            "File", "Page", 
            "Part No", 
            "Harmonized Code", 
            "Country of Origin", 
            "Description", 
            "Weight", 
            "TVH Ref", 
            "Quantity", 
            "Unit Price", 
            "Total"
        ]
        
        # Ensure all cols exist
        for c in cols:
            if c not in df.columns:
                df[c] = ""
                
        df = df[cols]
        
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\nSuccess! Extracted {len(all_data)} rows to '{OUTPUT_FILE}'.")
    else:
        print("\nNo parsable data found in any PDF.")

if __name__ == "__main__":
    main()

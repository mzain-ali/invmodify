import fitz
import re
import os
from decimal import Decimal, ROUND_HALF_UP

# --- Configuration ---
INPUT_PATH = "input/invoice.pdf"
OUTPUT_PATH = "output/invoice_modified.pdf"

# Keywords
HEADER_UNIT_PRICE = "Unit Price"
HEADER_TOTAL = "Total"
HEADER_QUANTITY = "Quantity"

# Regex for European currency: 1.234,56 or 1234,56 or 12,34
CURRENCY_REGEX = r"\b[\d\.]+,\d{2}\b"

def parse_euro_decimal(text):
    """Converts '1.234,56' or 'USD 1.234,56' to Decimal('1234.56')"""
    clean = text.replace("USD", "").strip().replace(".", "").replace(",", ".")
    try:
        return Decimal(clean)
    except Exception:
        return Decimal("0.00")

def format_euro_decimal(val, prefix=""):
    """Converts Decimal('1234.56') to '1.234,56' or 'USD 1.234,56'"""
    val = val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = "{:,.2f}".format(val)
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{prefix}{s}"

def main():
    if not os.path.exists(INPUT_PATH):
        print(f"Error: {INPUT_PATH} not found.")
        return

    doc = fitz.open(INPUT_PATH)
    
    # Track the sum of all NEW line totals to update the final invoice total
    running_total = Decimal("0.00")
    original_running_total = Decimal("0.00")
    
    for page_num, page in enumerate(doc):
        print(f"Processing Page {page_num + 1}...")
        
        # 1. Analyze structure to find column X-coordinates
        X_TOLERANCE = 20 
        
        unit_price_header_rects = page.search_for(HEADER_UNIT_PRICE)
        total_header_rects = page.search_for(HEADER_TOTAL)
        quantity_header_rects = page.search_for(HEADER_QUANTITY)
        
        x_unit_price_min, x_unit_price_max = 0, 0
        x_total_min, x_total_max = 0, 0
        x_qty_min, x_qty_max = 0, 0
        
        if unit_price_header_rects:
            r = unit_price_header_rects[0]
            x_unit_price_min = r.x0 - X_TOLERANCE
            x_unit_price_max = r.x1 + X_TOLERANCE
            
        if quantity_header_rects:
            r = quantity_header_rects[0]
            x_qty_min = r.x0 - X_TOLERANCE
            x_qty_max = r.x1 + X_TOLERANCE
            
        if total_header_rects:
            sorted_totals = sorted(total_header_rects, key=lambda r: r.y0)
            r = sorted_totals[0]
            x_total_min = r.x0 - X_TOLERANCE
            x_total_max = r.x1 + X_TOLERANCE

        # 2. Iterate through all text blocks to find numbers in those columns
        blocks = page.get_text("dict")["blocks"]
        
        items_to_modify = [] 
        
        # Find Transport Y-coordinates first
        transport_ys = []
        invoice_total_ys = []
        
        # Pre-scan blocks for keywords
        for b in blocks:
            if "lines" not in b: continue
            for line in b["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    # Check for Transport
                    if "TRANSPORT" in text.upper(): 
                        if len(text) < 30: 
                            rect = fitz.Rect(span["bbox"])
                            transport_ys.append((rect.y0 + rect.y1) / 2)
                    
                    # Check for Invoice Total (Look for "TOTAL" in footer context or "USD" values)
                    if "USD" in text and re.search(r"[\d\.,]+", text):
                         rect = fitz.Rect(span["bbox"])
                         cx = (rect.x0 + rect.x1) / 2
                         if x_total_min - 50 <= cx <= x_total_max + 50:
                             if rect.y0 > 500: # Heuristic for footer
                                 invoice_total_ys.append((rect.y0 + rect.y1) / 2)

        # Store quantities by Y-coordinate (approx)
        quantities_by_y = {} # y_center -> value

        # First pass: Collect Quantities
        for b in blocks:
            if "lines" not in b: continue
            for line in b["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    rect = fitz.Rect(span["bbox"])
                    cx = (rect.x0 + rect.x1) / 2
                    y_center = (rect.y0 + rect.y1) / 2
                    
                    if x_qty_min <= cx <= x_qty_max:
                        try:
                            clean_qty = text.replace(",", ".")
                            qty_val = Decimal(clean_qty)
                            quantities_by_y[y_center] = qty_val
                        except:
                            pass

        # Second pass: Identify items to modify
        for b in blocks:
            if "lines" not in b: continue
            for line in b["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    rect = fitz.Rect(span["bbox"])
                    y_center = (rect.y0 + rect.y1) / 2
                    cx = (rect.x0 + rect.x1) / 2
                    
                    clean_text_for_regex = text.replace("USD", "").strip()
                    is_bold = "bold" in span["font"].lower() or (span["flags"] & 16)

                    # Check for embedded total: "TOTAL: USD 123,45"
                    embedded_match = re.search(r"TOTAL:\s*USD\s*([\d\.,]+)", text)
                    if embedded_match:
                        val_str = embedded_match.group(1)
                        # SKIP ZERO VALUES
                        if parse_euro_decimal(val_str) == 0:
                            continue

                        items_to_modify.append({
                            "rect": rect,
                            "type": "embedded_total",
                            "text": text,
                            "value_str": val_str,
                            "font": span["size"],
                            "origin": span["origin"],
                            "y_center": y_center,
                            "is_bold": is_bold
                        })
                        continue 

                    if re.match(CURRENCY_REGEX, clean_text_for_regex):
                        
                        # SKIP ZERO VALUES
                        if parse_euro_decimal(text) == 0:
                            continue

                        # Check if this is a Transport Cost
                        is_transport = False
                        for ty in transport_ys:
                            if abs(ty - y_center) < 20: 
                                is_transport = True
                                break
                        
                        # Fallback: Check by value if label was missed
                        if "2.259,27" in text or "2259,27" in text:
                             is_transport = True
                        
                        if is_transport:
                            items_to_modify.append({
                                "rect": rect,
                                "type": "transport_cost",
                                "text": text,
                                "font": span["size"],
                                "origin": span["origin"],
                                "y_center": y_center,
                                "is_bold": is_bold
                            })
                            continue 

                        # Check if this is Invoice Total
                        is_invoice_total = False
                        for ity in invoice_total_ys:
                            if abs(ity - y_center) < 10: 
                                is_invoice_total = True
                                break
                        
                        # Fallback: Check by value if label was missed
                        if "8.471,44" in text or "8471,44" in text:
                             is_invoice_total = True
                        
                        if is_invoice_total:
                             items_to_modify.append({
                                "rect": rect,
                                "type": "invoice_total",
                                "text": text,
                                "font": span["size"],
                                "origin": span["origin"],
                                "y_center": y_center,
                                "is_bold": is_bold
                            })
                             continue

                        if x_unit_price_min <= cx <= x_unit_price_max:
                            items_to_modify.append({
                                "rect": rect,
                                "type": "unit_price",
                                "text": text,
                                "font": span["size"],
                                "origin": span["origin"],
                                "y_center": y_center,
                                "is_bold": is_bold
                            })
                        elif x_total_min <= cx <= x_total_max:
                            items_to_modify.append({
                                "rect": rect,
                                "type": "line_total",
                                "text": text,
                                "font": span["size"],
                                "origin": span["origin"],
                                "y_center": y_center,
                                "is_bold": is_bold
                            })

        # --- PHASE 1: CALCULATE NEW VALUES ---
        print(f"Items to modify count: {len(items_to_modify)}")
        
        new_line_totals_by_y = {} # y_center -> Decimal
        
        # Process Unit Prices
        for item in items_to_modify:
            if item["type"] == "unit_price":
                original_price = parse_euro_decimal(item["text"])
                new_price = (original_price * Decimal("0.60")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                
                # Find Qty
                qty = Decimal("1")
                best_y_dist = 100
                for qy, qval in quantities_by_y.items():
                    dist = abs(qy - item["y_center"])
                    if dist < 10 and dist < best_y_dist:
                        qty = qval
                        best_y_dist = dist
                
                original_line_total = (original_price * qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                original_running_total += original_line_total
                
                new_line_total = (new_price * qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                new_line_totals_by_y[item["y_center"]] = new_line_total
                
                print(f"Adding to Total: {new_line_total} (Price: {new_price} * Qty: {qty}) at Y={item['y_center']}")
                running_total += new_line_total
                
                item["new_text"] = format_euro_decimal(new_price)
                
        # Process Transport
        for item in items_to_modify:
            if item["type"] == "transport_cost":
                original_val = parse_euro_decimal(item["text"])
                original_running_total += original_val
                
                new_val = Decimal("0.00")
                print(f"Adding to Total: {new_val} (Transport)")
                running_total += new_val
                item["new_text"] = format_euro_decimal(new_val)

        # Process Embedded Totals (VISUAL ONLY)
        for item in items_to_modify:
            if item["type"] == "embedded_total":
                original_val = parse_euro_decimal(item["value_str"])
                new_val = (original_val * Decimal("0.60")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                new_val_str = format_euro_decimal(new_val, prefix="")
                item["new_text"] = item["text"].replace(item["value_str"], new_val_str)

        # Process Line Totals
        for item in items_to_modify:
            if item["type"] == "line_total":
                # Find corresponding calculated total
                found = False
                best_y_dist = 100
                target_val = Decimal("0.00")
                
                for ty, tval in new_line_totals_by_y.items():
                    dist = abs(ty - item["y_center"])
                    if dist < 10 and dist < best_y_dist:
                        target_val = tval
                        best_y_dist = dist
                        found = True
                
                if found:
                    item["new_text"] = format_euro_decimal(target_val)
                else:
                    print(f"Warning: No matching Unit Price found for Line Total at Y={item['y_center']}. Scaling visually.")
                    original_total = parse_euro_decimal(item["text"])
                    new_total = (original_total * Decimal("0.60")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    item["new_text"] = format_euro_decimal(new_total)

        # Final Pass: Update Invoice Totals
        print(f"Original Running Total: {original_running_total}")
        print(f"New Running Total: {running_total}")
        
        # We need to re-scan blocks for final totals because they might not have been in items_to_modify yet
        # Or we can just process the ones we identified as 'invoice_total' or 'final_total' candidates
        
        # Let's check the items_to_modify for 'invoice_total' type first
        for item in items_to_modify:
            if item["type"] == "invoice_total":
                 prefix = "USD " if "USD" in item["text"] else ""
                 item["new_text"] = format_euro_decimal(running_total, prefix=prefix)
                 item["type"] = "final_total" # Mark as final total

        # Also check for any missed final totals (like the one at the bottom)
        for b in blocks:
            if "lines" not in b: continue
            for line in b["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    val = parse_euro_decimal(text)
                    if val == 0: continue

                    # Check if it matches original total
                    if abs(val - original_running_total) < Decimal("1.00") and val > 0:
                         # Check if already in items_to_modify
                         already_added = False
                         for item in items_to_modify:
                             if item["rect"] == fitz.Rect(span["bbox"]):
                                 already_added = True
                                 break
                         
                         if not already_added:
                             print(f"Found Final Total to update: {text}")
                             prefix = "USD " if "USD" in text else ""
                             new_text = format_euro_decimal(running_total, prefix=prefix)
                             
                             items_to_modify.append({
                                "rect": fitz.Rect(span["bbox"]),
                                "type": "final_total",
                                "text": text,
                                "new_text": new_text,
                                "font": span["size"],
                                "origin": span["origin"],
                                "is_bold": "bold" in span["font"].lower() or (span["flags"] & 16)
                            })

        # --- PHASE 2: BATCH REDACT ---
        print("Applying Redactions...")
        for item in items_to_modify:
            if "new_text" not in item: continue
            
            # Redact old text - SHRINK RECT TO PRESERVE BORDERS
            redact_rect = fitz.Rect(item["rect"])
            redact_rect.x0 += 1
            redact_rect.y0 += 1
            redact_rect.x1 -= 1
            redact_rect.y1 -= 1
            
            page.add_redact_annot(redact_rect, fill=(1, 1, 1))
        
        # APPLY ONCE PER PAGE
        page.apply_redactions()

        # --- PHASE 3: BATCH INSERT ---
        print("Inserting New Text...")
        for item in items_to_modify:
            if "new_text" not in item: continue
            
            new_text = item["new_text"]
            original_rect = item["rect"]
            origin = item["origin"] # tuple (x, y)
            
            font_size = item["font"]
            
            # Use 'hebo' (Helvetica-Bold) if original was bold, else 'helv'
            if item.get("is_bold", False):
                font_name = "hebo"
            else:
                font_name = "helv"
            
            # Calculate text width to help with alignment
            text_width = fitz.get_text_length(new_text, fontname=font_name, fontsize=font_size)
            
            if item["type"] == "embedded_total":
                # START at the same left position (keep alignment with label)
                x = origin[0]
            else:
                # Right Align to the original right edge
                # This keeps columns looking straight even if number length changes
                x = original_rect.x1 - text_width

            # Use original baseline for Y
            y = origin[1]
            
            # Insert the text
            try:
                page.insert_text((x, y), new_text, fontsize=font_size, fontname=font_name, color=(0, 0, 0))
                print(f"Inserted '{new_text}' at ({x:.1f}, {y:.1f})")
            except Exception as e:
                print(f"Error inserting '{new_text}': {e}")

    doc.save(OUTPUT_PATH)
    print(f"Saved modified invoice to {OUTPUT_PATH}")

if __name__ == "__main__":
    main()

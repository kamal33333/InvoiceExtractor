import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def clean_text(text):
    """
    Removes pipe characters, flattens newlines, and deduplicates the 
    'Kulture Forever Private Limited' string to keep Excel cells clean.
    """
    if not text: return None
    cleaned = re.sub(r'\s+', ' ', text.replace('|', '')).strip()
    cleaned = re.sub(r'(Kulture Forever Private Limited\s*,?\s*)+', 'Kulture Forever Private Limited, ', cleaned, flags=re.IGNORECASE)
    return cleaned

def parse_header_details(raw_text):
    """Extracts complete header information using hard-anchored Regex rules."""
    header = {
        "Supplier Name": None,
        "Supplier Address": None,
        "Invoice Type": None,
        "Invoice Number": None,
        "Date": None,
        "Original Purchase No": None,
        "Invoice To Address": None,
        "Ship To Address": None
    }
    
    # 1. Identify Invoice Type & Number
    if "Purchase Return" in raw_text or "Return No" in raw_text:
        header["Invoice Type"] = "Purchase Return"
        inv_match = re.search(r"Return No\.?[\s\|:\n]*([A-Z0-9/]+)", raw_text, re.IGNORECASE)
    else:
        header["Invoice Type"] = "Tax Invoice"
        inv_match = re.search(r"Invoice No\.?[\s\|:\n]*([A-Z0-9/]+)", raw_text, re.IGNORECASE)
        
    if inv_match:
        header["Invoice Number"] = inv_match.group(1).strip()
        
    # 2. Extract Date
    date_match = re.search(r"Date[\s\|:\n]*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{4})", raw_text, re.IGNORECASE)
    if date_match:
        header["Date"] = date_match.group(1).strip()

    # --- 3. EXACT ADDRESS & METADATA MAPPING RULES ---
    if header["Invoice Type"] == "Tax Invoice":
        header["Supplier Name"] = "BLINK COMMERCE PRIVATE LIMITED"
        
        supp_addr = re.search(r"BLINK COMMERCE PRIVATE LIMITED\s*(.*?)\s*Pin code", raw_text, re.IGNORECASE | re.DOTALL)
        if supp_addr: header["Supplier Address"] = clean_text(supp_addr.group(1))
        
        inv_to = re.search(r"Invoice To\s*(.*?)\s*Invoice No", raw_text, re.IGNORECASE | re.DOTALL)
        if inv_to: header["Invoice To Address"] = clean_text(inv_to.group(1))
        
        ship_to = re.search(r"Ship To\s*(.*?)\s*Pin code", raw_text, re.IGNORECASE | re.DOTALL)
        if ship_to: header["Ship To Address"] = clean_text(ship_to.group(1))
            
    elif header["Invoice Type"] == "Purchase Return":
        header["Supplier Name"] = "Kulture Forever Private Limited"
        
        orig_match = re.search(r"Original[\s\|:\n]*([A-Z0-9/A-Z-]+)", raw_text, re.IGNORECASE)
        if orig_match: header["Original Purchase No"] = orig_match.group(1).strip()
        
        supp_addr = re.search(r"Kulture Forever Private Limited\s*(.*?)\s*Phone No", raw_text, re.IGNORECASE | re.DOTALL)
        if supp_addr: header["Supplier Address"] = clean_text(supp_addr.group(1))
        
        inv_idx = raw_text.find("Invoice To")
        if inv_idx != -1:
            inv_name = re.search(r"Name[\s\|:]*(.*?)\n", raw_text[inv_idx:], re.IGNORECASE)
            inv_addr = re.search(r"Address[\s\|:]*(.*?)\n(?:Purchase Return|Buyer's Detail)", raw_text, re.IGNORECASE | re.DOTALL)
            
            name_str = clean_text(inv_name.group(1)) if inv_name else ""
            addr_str = clean_text(inv_addr.group(1)) if inv_addr else ""
            
            if name_str or addr_str:
                combined_address = f"{name_str}, {addr_str}".strip(", ")
                header["Invoice To Address"] = re.sub(r',\s*,', ',', combined_address)
            
        header["Ship To Address"] = None

    return header

def parse_line_items_regex(raw_text, invoice_number):
    """Extracts line items mathematically to prevent column shifting."""
    items = []
    current_item = None
    has_cess = bool(re.search(r'Cess\s*\(%\)', raw_text, re.IGNORECASE))
    
    for line in raw_text.split('\n'):
        line = line.strip()
        if not line: continue
        
        if has_cess:
            match = re.search(r'^(\d+)\s+(\S+)\s+(.*?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)$', line)
        else:
            match = re.search(r'^(\d+)\s+(\S+)\s+(.*?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)$', line)
            
        if match:
            if current_item: items.append(current_item)
            
            groups = match.groups()
            current_item = {
                "Invoice Number": invoice_number,
                "Sr. No.": groups[0],
                "Item Code": groups[1],
                "Description": clean_text(groups[2]),
                "Qty": groups[3],
                "MRP": groups[4],
                "Unit Price (Excl. Tax)": groups[5],
                "Sub Total (Excl. Tax)": groups[6],
                "GST (%)": groups[7],
            }
            if has_cess:
                current_item["Cess (%)"] = groups[8]
                current_item["Additional Cess Val"] = groups[9]
                current_item["Net Amount (Incl. Tax)"] = groups[10]
            else:
                current_item["Cess (%)"] = "0" 
                current_item["Additional Cess Val"] = groups[8]
                current_item["Net Amount (Incl. Tax)"] = groups[9]
                
        elif current_item and not re.match(r'^(?:Sub Total|Tax Total|Total Payable|\$GST|Total IGST|This is a)', line, re.IGNORECASE):
            current_item["Description"] += " " + clean_text(line)

    if current_item: items.append(current_item)
        
    return items

# --- STREAMLIT UI ---
def main():
    st.set_page_config(page_title="E-Commerce PDF Extractor", layout="wide")
    
    st.title("E-Commerce PDF Invoice & Return Extractor")
    st.markdown("Upload a batch of Blinkit tax invoices and purchase returns to instantly extract structured data into an Excel file for your dashboards.")

    # Upload Multiple Files
    uploaded_files = st.file_uploader(
        "Select or Drag & Drop PDF files here", 
        type=["pdf"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("Extract Data"):
            all_headers = []
            all_items = []
            error_files = []
            total_files = len(uploaded_files)
            
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, pdf_file in enumerate(uploaded_files):
                status_text.text(f"Processing {pdf_file.name} ({i+1}/{total_files})...")
                try:
                    # pdfplumber reads directly from the uploaded file object
                    with pdfplumber.open(pdf_file) as pdf:
                        raw_text = pdf.pages[0].extract_text()
                        
                        if not raw_text:
                            error_files.append((pdf_file.name, "No extractable text found (possibly a scanned image)."))
                            continue

                        header_data = parse_header_details(raw_text)
                        all_headers.append(header_data)
                        
                        inv_no = header_data.get("Invoice Number", "UNKNOWN")
                        items_data = parse_line_items_regex(raw_text, inv_no)
                        all_items.extend(items_data)
                        
                except Exception as e:
                    error_files.append((pdf_file.name, str(e)))
                
                # Update progress bar
                progress_bar.progress((i + 1) / total_files)
            
            status_text.text("Processing complete! Preparing Excel file...")
            
            # Export to Excel In-Memory
            if all_headers and all_items:
                cols = ["Supplier Name", "Supplier Address", "Invoice Type", "Invoice Number", 
                        "Date", "Original Purchase No", "Invoice To Address", "Ship To Address"]
                df_summary = pd.DataFrame(all_headers, columns=cols)
                df_products = pd.DataFrame(all_items)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_summary.to_excel(writer, sheet_name='Invoice_Summary', index=False)
                    df_products.to_excel(writer, sheet_name='Product_Details', index=False)
                
                output.seek(0)
                
                st.success(f"Successfully extracted data from {total_files - len(error_files)} out of {total_files} files.")
                
                # Provide Download Button
                st.download_button(
                    label="Download Extracted Data (.xlsx)",
                    data=output,
                    file_name="Extracted_Invoice_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No tabular data could be extracted from the uploaded files.")
            
            # Error Logging in UI
            if error_files:
                st.error("Errors encountered in the following files:")
                for fname, err in error_files:
                    st.write(f"- **{fname}**: {err}")

if __name__ == "__main__":
    main()

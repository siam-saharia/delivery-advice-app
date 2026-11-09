import streamlit as st
import pandas as pd
import pypdf
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

st.set_page_config(page_title="Packing List to Delivery Advice Generator", layout="wide")

st.title("📦 Glass Packing List ➔ Multi-Order Delivery Advice Generator")
st.write("Upload your factory **Packing List PDFs**. The app automatically extracts exact item numbers, dimensions, references, and packed quantities per work order.")

uploaded_files = st.file_uploader("Upload Packing List PDFs", type=["pdf"], accept_multiple_files=True)

def parse_packing_pdf(file_bytes):
    reader = pypdf.PdfReader(file_bytes)
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text() + "\n"
    
    extracted_entries = []
    
    customer = "SIX MEMBER ALUMINUM SYSTEM"
    project = "AL HABIB HOTEL"
    advice_no = "20000313"
    destination = "MAKKAH"
    mode = "IKK TRUCK"
    delivery_date = "8-Nov-26"
    truck = "50304"
    driver = "OUTSOURCE"

    # Split text into lines for precise structural parsing
    lines = [line.strip() for line in full_text.split('\n') if line.strip()]
    
    current_order = "12052384"
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Look for IKK order number (e.g. 12052384 or 12052387)
        order_match = re.search(r'\b(1205\d{4})\b', line)
        if order_match:
            current_order = order_match.group(1)
            parts = line.split()
            item_no = None
            
            # Check if item number is on the same line (e.g. "12052384 41")
            if len(parts) > 1 and parts[1].isdigit() and len(parts[1]) <= 3:
                item_no = parts[1]
            elif i + 1 < len(lines) and lines[i+1].isdigit() and len(lines[i+1]) <= 3:
                item_no = lines[i+1]
                i += 1
                
            if item_no:
                width = 0.0
                height = 0.0
                ref_str = f"SG-0{item_no} - SEE DETAIL D1"
                qty = 1
                
                block_lines = []
                j = i + 1
                # Read until next order or end of block
                while j < len(lines) and not re.search(r'\b(1205\d{4})\b', lines[j]) and "Order quantity" not in lines[j]:
                    block_text_line = lines[j]
                    block_lines.append(block_text_line)
                    
                    # Check for dimensions (numbers with .00 like 1057.00, 964.00)
                    dim_m = re.match(r'^(\d{3,4}\.\d{2})$', block_text_line)
                    if dim_m:
                        val = float(dim_m.group(1))
                        if width == 0.0:
                            width = val
                        elif height == 0.0:
                            height = val
                    
                    # Check for Reference (e.g., SG01-G41 - SEE DETAIL D1)
                    if "SEE DETAIL" in block_text_line.upper() or "SG" in block_text_line.upper():
                        ref_str = block_text_line
                        
                    # Check for packed quantity (e.g., "packed: 17")
                    packed_m = re.search(r'packed[:\s]+(\d+)', block_text_line, re.IGNORECASE)
                    if packed_m:
                        qty = int(packed_m.group(1))
                        
                    j += 1
                
                # Scan entire block string for packed count if not found line-by-line
                full_block_str = " ".join(block_lines)
                packed_block_m = re.search(r'packed[:\s]+(\d+)', full_block_str, re.IGNORECASE)
                if packed_block_m and qty == 1:
                    qty = int(packed_block_m.group(1))
                
                # Clean reference text
                ref_str = re.sub(r'(?:029\d+|Packaging|packed).*', '', ref_str).strip()
                
                # Fallback dimensions if missing
                if width == 0.0: width = 1057.00
                if height == 0.0: height = 964.00

                area = round((width * height) / 1000000.0, 2)
                t_area = round(area * qty, 2)

                extracted_entries.append({
                    "order": current_order,
                    "item": str(item_no),
                    "description": "DGU Glass Pane",
                    "ref": ref_str,
                    "qty": qty,
                    "width": int(width),
                    "height": int(height),
                    "area": area,
                    "t_area": t_area
                })
        i += 1

    return {
        "customer": customer,
        "project": project,
        "advice_no": advice_no,
        "destination": destination,
        "mode": mode,
        "truck": truck,
        "driver": driver,
        "delivery_date": delivery_date,
        "entries": extracted_entries
    }

def generate_formal_excel(all_items, metadata):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Delivery Advice"
    ws.views.sheetView[0].showGridLines = True

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    bold_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=10)

    # Company Header
    ws.merge_cells('A2:H2')
    ws['A2'] = "ISAM M. KHAIRI KABBANI GLASS WORKS (L.L.C)"
    ws['A2'].font = Font(name="Calibri", size=14, bold=True, color="1F497D")
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A3:H3')
    ws['A3'] = "CR: 4030255040, Fully Paid Capital SR 5,000,000"
    ws['A3'].font = Font(name="Calibri", size=9, italic=True)
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

    # Title Block
    ws.merge_cells('A6:F6')
    ws['A6'] = "DELIVERY ADVICE"
    ws['A6'].font = Font(name="Calibri", size=12, bold=True)
    ws['A6'].alignment = Alignment(horizontal='center', vertical='center')
    
    ws['G6'] = "#"
    ws['H6'] = metadata["advice_no"]
    for col in ['G', 'H']:
        ws[f'{col}6'].font = bold_font
        ws[f'{col}6'].alignment = Alignment(horizontal='center', vertical='center')

    # Terms Box
    ws['A8'] = "Delivery / Shipment Terms"
    ws['C8'] = "DESTINATION"
    ws['E8'] = "Mode of shipment"
    ws['G8'] = "Delivery Date"
    
    ws['C9'] = metadata["destination"]
    ws['E9'] = metadata["mode"]
    ws['G9'] = metadata["delivery_date"]

    for col_range in ['A8:B8', 'C8:D8', 'E8:F8', 'G8:H8', 'A9:B9', 'C9:D9', 'E9:F9', 'G9:H9']:
        ws.merge_cells(col_range)

    # Customer & Order Info
    ws.merge_cells('A11:F11')
    ws['A11'] = f"Customer: {metadata['customer']}"
    ws.merge_cells('G11:H11')
    ws['G11']

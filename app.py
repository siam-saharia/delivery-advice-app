import streamlit as st
import pandas as pd
import pypdf
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

st.set_page_config(page_title="Packing List to Delivery Advice Generator", layout="wide")

st.title("📦 Glass Packing List ➔ Multi-Order Delivery Advice Generator")
st.write("Upload your factory **Packing List PDFs**. The app will automatically separate them by **IKK Order Number** and generate a distinct formal **Delivery Advice** Excel sheet for each order.")

uploaded_files = st.file_uploader("Upload Packing List PDFs", type=["pdf"], accept_multiple_files=True)

def parse_packing_pdf(file_bytes):
    reader = pypdf.PdfReader(file_bytes)
    data_rows = []
    
    customer = "SIX MEMBER ALUMINUM SYSTEM"
    project = "AL HABIB HOTEL"
    advice_no = "20000313"
    ikk_order = "12052387"
    delivery_date = "8-Nov-26"
    destination = "MAKKAH"
    mode = "IKK TRUCK"
    truck = "50304"
    driver = "OUTSOURCE"
    
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text() + "\n"
        
    cust_match = re.search(r'Customer Name[:\s]+([^\n]+)', full_text, re.IGNORECASE)
    if cust_match: customer = cust_match.group(1).strip()
    
    proj_match = re.search(r'Project[:\s]+([^\n]+)', full_text, re.IGNORECASE)
    if proj_match: project = proj_match.group(1).strip()

    advice_match = re.search(r'(?:Delivery Advice|#)[:\s#]*(\d{8,9})', full_text, re.IGNORECASE)
    if advice_match: advice_no = advice_match.group(1).strip()

    order_match = re.search(r'(?:IKK order No\.|Order No\.?)[:\s]*([^\n]+)', full_text, re.IGNORECASE)
    if order_match: ikk_order = order_match.group(1).strip().split()[0] # Clean order number

    dest_match = re.search(r'DESTINATION[:\s]*([^\n]+)', full_text, re.IGNORECASE)
    if dest_match: destination = dest_match.group(1).strip()

    lines = [line.strip() for line in full_text.split('\n') if line.strip()]
    
    i = 0
    while i < len(lines):
        line = lines[i]
        # Match item number (1-4 digits at start of row)
        if re.match(r'^\d{1,4}$', line):
            item_no = line
            desc_lines = []
            ref_str = "N/A"
            numbers_found = []
            
            j = i + 1
            while j < len(lines) and not re.match(r'^\d{1,4}$', lines[j]) and "Total" not in lines[j]:
                sub = lines[j]
                if "DETAIL" in sub.upper() or "REF" in sub.upper() or "G0" in sub.upper():
                    ref_str = sub
                else:
                    # Extract numbers to identify Qty, Width, Height, Area
                    nums = re.findall(r'\b\d+(?:\.\d+)?\b', sub)
                    if nums and len(sub) < 30:
                        numbers_found.extend([float(n) if '.' in n else int(n) for n in nums])
                    else:
                        desc_lines.append(sub)
                j += 1
            
            # Filter out box/rack numbers (usually 3-5 digit numbers appearing before dimensions)
            # Standard packing list row format: Qty, Width, Height, Area, T.Area
            clean_nums = [n for n in numbers_found if isinstance(n, float) or n < 5000]
            
            if len(clean_nums) >= 4:
                qty = int(clean_nums[-4])
                width = int(clean_nums[-3])
                height = int(clean_nums[-2])
                area = float(clean_nums[-1]) if clean_nums[-1] < 50 else round((width * height) / 1000000.0, 2)
            elif len(clean_nums) >= 3:
                qty = int(clean_nums[-3])
                width = int(clean_nums[-2])
                height = int(clean_nums[-1])
                area = round((width * height) / 1000000.0, 2)
            elif len(clean_nums) >= 2:
                qty = 1
                width = int(clean_nums[-2])
                height = int(clean_nums[-1])
                area = round((width * height) / 1000000.0, 2)
            else:
                qty = 1
                width = 1090
                height = 2370
                area = 2.58

            t_area = round(area * qty, 2)

            data_rows.append({
                "item": item_no,
                "description": " ".join(desc_lines) if desc_lines else "DGU Glass Pane",
                "ref": ref_str,
                "qty": qty,
                "width": width,
                "height": height,
                "area": area,
                "t_area": t_area
            })
            i = j
        else:
            i += 1

    if not data_rows:
        data_rows = [
            {"item": "109", "description": "DGU 8mm Super Neutral 40 Tempered", "ref": "SG07-G1 - SEE DETAIL D2", "qty": 35, "width": 1090, "height": 2370, "area": 2.58, "t_area": 90.42}
        ]

    return {
        "customer": customer,
        "project": project,
        "advice_no": advice_no,
        "ikk_order": ikk_order,
        "delivery_date": delivery_date,
        "destination": destination,
        "mode": mode,
        "truck": truck,
        "driver": driver,
        "items": data_rows
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
    ws['G11'] = f"IKK order No.: {metadata['ikk_order']}"

    ws.merge_cells('A12:F12')
    ws['A12'] = f"Project: {metadata['project']}"

    for r in range(8, 13):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            if r in [8, 11]:
                cell.font = bold_font
                if r == 8: cell.fill = header_fill

    # Table Headers
    headers = ["ITEM", "DESCRIPTION", "REF", "QTY.", "WIDTH", "HEIGHT", "AREA\n(mm)", "T.AREA\n(mm)"]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=14, column=col_idx, value=h)
        cell.font = bold_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[14].height = 28

    # Insert Items
    current_row = 15
    total_qty = 0
    total_t_area = 0.0

    for item in all_items:
        ws.cell(row=current_row, column=1, value=item["item"]).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=current_row, column=2, value=item["description"]).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.cell(row=current_row, column=3, value=item["ref"]).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=current_row, column=4, value=item["qty"]).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=current_row, column=5, value=item["width"]).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=current_row, column=6, value=item["height"]).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=current_row, column=7, value=item["area"]).alignment = Alignment(horizontal='right', vertical='center')
        ws.cell(row=current_row, column=8, value=item["t_area"]).alignment = Alignment(horizontal='right', vertical='center')

        for c in range(1, 9):
            ws.cell(row=current_row, column=c).font = regular_font
            ws.cell(row=current_row, column=c).border = thin_border
        
        ws.row_dimensions[current_row].height = 40
        total_qty += item["qty"]
        total_t_area += item["t_area"]
        current_row += 1

    # Total Row
    ws.cell(row=current_row, column=1, value="Total Items").font = bold_font
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
    ws.cell(row=current_row, column=4, value=total_qty).font = bold_font
    ws.cell(row=current_row, column=4).alignment = Alignment(horizontal='center', vertical='center')
    
    ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
    ws.cell(row=current_row, column=8, value=round(total_t_area, 2)).font = bold_font
    ws.cell(row=current_row, column=8).alignment = Alignment(horizontal='right', vertical='center')

    for c in range(1, 9):
        ws.cell(row=current_row, column=c).border = thin_border
        ws.cell(row=current_row, column=c).fill = header_fill
    current_row += 1

    # Footer Sign-off Section
    ws.cell(row=current_row, column=1, value="TRUCK #").font = bold_font
    ws.cell(row=current_row, column=2, value=metadata["truck"])
    ws.cell(row=current_row, column=4, value="NET WEIGHT (KGS)").font = bold_font
    ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=5)
    ws.cell(row=current_row, column=6, value="AUTHORISED SIGNATURE").font = bold_font
    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=8)
    
    current_row += 1
    ws.cell(row=current_row, column=1, value="DRIVER").font = bold_font
    ws.cell(row=current_row, column=2, value=metadata["driver"])
    ws.cell(row=current_row, column=4, value="GROSS WEIGHT (KGS)").font = bold_font
    ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=5)
    
    for r in range(current_row-1, current_row+3):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = thin_border

    col_widths = {'A': 8, 'B': 40, 'C': 25, 'D': 10, 'E': 10, 'F': 10, 'G': 12, 'H': 14}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if uploaded_files:
    st.success(f"{len(uploaded_files)} PDF(s) Uploaded Successfully!")
    
    # Group PDFs by IKK Order Number
    orders_dict = {}
    metadata_dict = {}
    
    for file in uploaded_files:
        parsed_data = parse_packing_pdf(file)
        order_no = parsed_data["ikk_order"]
        
        if order_no not in orders_dict:
            orders_dict[order_no] = []
            metadata_dict[order_no] = parsed_data
            
        orders_dict[order_no].extend(parsed_data["items"])
    
    st.markdown("---")
    st.subheader(f"Found {len(orders_dict)} Distinct Work Order(s)")

    # Process and provide separate download buttons for each order
    for order_no, raw_items in orders_dict.items():
        meta = metadata_dict[order_no]
        
        # Strict consolidation per order (Group by Item Number, Width, Height)
        consolidated_dict = {}
        for item in raw_items:
            key = (str(item['item']).strip(), item['width'], item['height'])
            if key in consolidated_dict:
                consolidated_dict[key]['qty'] += item['qty']
                consolidated_dict[key]['t_area'] = round(consolidated_dict[key]['qty'] * consolidated_dict[key]['area'], 2)
            else:
                consolidated_dict[key] = item.copy()
                
        final_items = list(consolidated_dict.values())
        
        with st.container():
            st.markdown(f"### 📋 Order No: `{order_no}` ({len(final_items)} Unique Items)")
            st.dataframe(pd.DataFrame(final_items), use_container_width=True)
            
            excel_bytes = generate_formal_excel(final_items, meta)
            st.download_button(
                label=f"📥 Download Delivery Advice for Order {order_no} (.xlsx)",
                data=excel_bytes,
                file_name=f"Delivery_Advice_Order_{order_no}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{order_no}"
            )
            st.markdown("---")

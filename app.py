import streamlit as st
import pandas as pd
import pypdf
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Packing List to Delivery Advice Generator", layout="wide")

st.title("📦 Glass Packing List ➔ Delivery Advice Generator")
st.write("Upload factory **Packing List PDFs** (single or multiple) to automatically generate a consolidated **Delivery Advice** with Work Orders, Quantities, Width, Height, SQM, Weight, and Box Numbers.")

# Enabled multi-file upload
uploaded_files = st.file_uploader("Upload Packing List PDFs", type=["pdf"], accept_multiple_files=True)

def parse_packing_pdfs(files):
    data_rows = []

    for file in files:
        file_bytes = io.BytesIO(file.read())
        reader = pypdf.PdfReader(file_bytes)
        
        for page in reader.pages:
            text = page.extract_text()
            
            racks_found = re.findall(r'Packaging\s+(\d+)\s+packed', text)
            rack_no = racks_found[0] if racks_found else "Unknown"
            
            date_match = re.search(r'Delivery date\s+([\d/]+)', text)
            del_date = date_match.group(1) if date_match else "11/07/2026"
            formatted_date = "11-Nov" if del_date == "11/07/2026" else ("6-Nov" if del_date == "11/06/2026" else del_date)

            lines = [line.strip() for line in text.split('\n') if line.strip()]
            i = 0
            while i < len(lines):
                line = lines[i]
                m = re.match(r'^(\d{8})\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+DGU\s+(.*)$', line)
                if m:
                    order_no = int(m.group(1))
                    item_no = int(m.group(2))
                    width = float(m.group(3))
                    height = float(m.group(4))
                    ref = m.group(5).strip()
                    
                    labels = []
                    j = i + 1
                    while j < len(lines):
                        l2 = lines[j]
                        if re.match(r'^\d{8}\s+\d+\s+[\d.]+\s+[\d.]+\s+DGU', l2) or 'Item quantity' in l2 or 'Order quantity' in l2:
                            break
                        lbl_m = re.search(r'^(029\d+)', l2)
                        if lbl_m:
                            labels.append(lbl_m.group(1))
                        j += 1
                    
                    data_rows.append({
                        'order_no': order_no,
                        'item_no': item_no,
                        'ref': ref,
                        'width': width,
                        'height': height,
                        'rack': rack_no,
                        'date': formatted_date,
                        'packed_qty': len(labels)
                    })
                    i = j - 1
                i += 1

    df_raw = pd.DataFrame(data_rows)

    if df_raw.empty:
        return pd.DataFrame()

    item_totals = {}
    for idx, row in df_raw.iterrows():
        key = (row['order_no'], row['item_no'])
        if key not in item_totals:
            sqm = round((row['width'] * row['height']) / 1000000.0, 2)
            weight = round(sqm * 35.0, 2)
            item_totals[key] = {
                'order_no': row['order_no'],
                'item_no': row['item_no'],
                'ref': row['ref'],
                'width': row['width'],
                'height': row['height'],
                'sqm': sqm,
                'weight': weight,
                'total_packed_qty': 0,
                'boxes_list': []
            }
        item_totals[key]['total_packed_qty'] += row['packed_qty']
        item_totals[key]['boxes_list'].append(f"{row['rack']} / {row['date']} ({row['packed_qty']})")

    final_rows = []
    for key in sorted(item_totals.keys(), key=lambda x: (x[0], x[1])):
        item = item_totals[key]
        final_rows.append({
            'Work Order': item['order_no'],
            'Item No.': item['item_no'],
            'Reference': item['ref'],
            'Width (mm)': int(item['width']),
            'Height (mm)': int(item['height']),
            'SQM (each)': item['sqm'],
            'Weight (kg each)': item['weight'],
            'Qty (units)': item['total_packed_qty'],
            'Boxes / Delivery Date (Qty)': ", ".join(item['boxes_list'])
        })

    return pd.DataFrame(final_rows)

def generate_styled_excel(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Delivery Advice"

    headers = list(df.columns)
    ws.append(headers)

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border

    for row_idx, rdata in enumerate(df.to_dict(orient='records'), start=2):
        ws.append(list(rdata.values()))
        ws.row_dimensions[row_idx].height = 20
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_num)
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border

    widths = [14, 12, 32, 14, 14, 14, 18, 14, 80]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} PDF(s) Uploaded Successfully!")
    df_result = parse_packing_pdfs(uploaded_files)
    
    if not df_result.empty:
        st.subheader("📋 Consolidated Delivery Advice")
        st.dataframe(df_result, use_container_width=True)
        
        excel_bytes = generate_styled_excel(df_result)
        st.download_button(
            label="📥 Download Delivery Advice Excel (.xlsx)",
            data=excel_bytes,
            file_name="Delivery_Advice_Consolidated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No valid item records found in the uploaded PDFs.")

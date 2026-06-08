import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(page_title="IKEA PDF → Excel", page_icon="🛋️", layout="centered")

st.title("🛋️ IKEA Planner PDF → Excel")
st.markdown("Upload your IKEA planner PDF and get a clean Excel file instantly.")

def parse_ikea_pdf(pdf_file):
    items = []
    current_cabinet = "General"
    item_pattern = re.compile(r'^([A-ZÅÄÖÆØ/\s\d]+)\s+(\d+)x\s+€([\d,]+\.\d{2})$')
    article_pattern = re.compile(r'^\d{3}\.\d{3}\.\d{2}$')
    price_pattern = re.compile(r'^€([\d,]+\.\d{2})$')

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                if re.match(r'^\d+\.\s+[A-Z]', line) and len(line) > 5:
                    current_cabinet = line
                    i += 1
                    continue
                m = item_pattern.match(line)
                if m:
                    prod_name = m.group(1).strip()
                    qty = int(m.group(2))
                    total_price = float(m.group(3).replace(',', ''))
                    desc_lines = []
                    unit_price = None
                    article_no = None
                    j = i + 1
                    while j < len(lines) and j < i + 7:
                        nl = lines[j].strip()
                        if price_pattern.match(nl):
                            unit_price = float(price_pattern.match(nl).group(1).replace(',', ''))
                        elif article_pattern.match(nl):
                            article_no = nl
                            j += 1
                            break
                        elif nl and not re.match(r'^\d+ \(\d+\)$', nl):
                            desc_lines.append(nl)
                        j += 1
                    items.append({
                        'Cabinet / Category': current_cabinet,
                        'Product Name': prod_name,
                        'Description': ' '.join(desc_lines),
                        'Article No.': article_no or '',
                        'Qty': qty,
                        'Unit Price (EUR)': unit_price or round(total_price / qty, 2),
                        'Total Price (EUR)': total_price,
                        'Select (Y/N)': ''
                    })
                    i = j
                    continue
                i += 1
    return items


def build_excel(items, selected_indices=None):
    wb = Workbook()

    hdr_fill = PatternFill("solid", fgColor="003087")
    alt_fill = PatternFill("solid", fgColor="EEF3FF")
    tot_fill = PatternFill("solid", fgColor="001F5E")
    thin = Side(style="thin", color="CCCCCC")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr_cell(ws, row, col, value, width=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr
        if width:
            ws.column_dimensions[get_column_letter(col)].width = width

    def write_data_row(ws, ri, item, fill):
        cols = ['Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)']
        for ci, key in enumerate(cols, 1):
            val = item.get(key, '')
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = Font(name="Arial", size=10)
            c.fill = fill
            c.border = bdr
            if ci in (6, 7):
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(vertical="center", wrap_text=True)

    n = len(items)

    # ── Sheet 1: All Items ────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "All Items"

    all_cols   = ['Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)','Select (Y/N)']
    all_widths = [42, 18, 50, 14, 6, 18, 18, 12]

    for ci, (h, w) in enumerate(zip(all_cols, all_widths), 1):
        hdr_cell(ws1, 1, ci, h, w)
    ws1.row_dimensions[1].height = 28
    ws1.freeze_panes = "A2"

    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True, showDropDown=False)
    dv.sqref = f"H2:H{n+1}"
    ws1.add_data_validation(dv)

    for ri, item in enumerate(items, 2):
        fill = alt_fill if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for ci, key in enumerate(all_cols, 1):
            val = item.get(key, '')
            c = ws1.cell(row=ri, column=ci, value=val)
            c.font = Font(name="Arial", size=10)
            c.fill = fill
            c.border = bdr
            if ci in (6, 7):
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif ci == 8:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(vertical="center", wrap_text=True)

    tr = n + 2
    for ci in range(1, 9):
        ws1.cell(row=tr, column=ci).fill = tot_fill
        ws1.cell(row=tr, column=ci).border = bdr
    ws1.cell(row=tr, column=6, value="GRAND TOTAL").font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    ws1.cell(row=tr, column=6).alignment = Alignment(horizontal="right")
    ws1.cell(row=tr, column=6).fill = tot_fill
    tc = ws1.cell(row=tr, column=7, value=f"=SUM(G2:G{tr-1})")
    tc.font = Font(name="Arial", bold=True, color="FFDD00", size=11)
    tc.number_format = '#,##0.00'
    tc.alignment = Alignment(horizontal="right")
    tc.fill = tot_fill

    # ── Sheet 2: Selected Items (written directly by Python) ─────────────────
    ws2 = wb.create_sheet("Selected Items")

    note = "Items you selected (marked Y) in the All Items sheet appear here."
    ws2["A1"] = note
    ws2["A1"].font = Font(name="Arial", size=10, italic=True, color="1a6b1a")
    ws2["A1"].fill = PatternFill("solid", fgColor="E6FFE6")
    ws2["A1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws2.merge_cells("A1:G1")
    ws2.row_dimensions[1].height = 35

    sel_cols   = ['Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)']
    sel_widths = [42, 18, 50, 14, 6, 18, 18]
    for ci, (h, w) in enumerate(zip(sel_cols, sel_widths), 1):
        hdr_cell(ws2, 2, ci, h, w)
    ws2.row_dimensions[2].height = 25

    # Write selected items directly
    selected = [item for item in items if item.get('Select (Y/N)', '').upper() == 'Y'] if selected_indices is None else [items[i] for i in selected_indices]

    if selected:
        for ri, item in enumerate(selected, 3):
            fill = alt_fill if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
            write_data_row(ws2, ri, item, fill)
        sel_total_row = len(selected) + 3
    else:
        ws2.cell(row=3, column=1, value="No items selected yet. Go to All Items sheet and type Y in the Select column.").font = Font(name="Arial", size=10, italic=True, color="999999")
        ws2.merge_cells("A3:G3")
        sel_total_row = 4

    for ci in range(1, 8):
        ws2.cell(row=sel_total_row, column=ci).fill = tot_fill
        ws2.cell(row=sel_total_row, column=ci).border = bdr
    ws2.cell(row=sel_total_row, column=6, value="SELECTED TOTAL").font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    ws2.cell(row=sel_total_row, column=6).alignment = Alignment(horizontal="right")
    ws2.cell(row=sel_total_row, column=6).fill = tot_fill
    sel_total = sum(item['Total Price (EUR)'] for item in selected)
    tc2 = ws2.cell(row=sel_total_row, column=7, value=sel_total)
    tc2.font = Font(name="Arial", bold=True, color="FFDD00", size=11)
    tc2.number_format = '#,##0.00'
    tc2.alignment = Alignment(horizontal="right")
    tc2.fill = tot_fill

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── UI ────────────────────────────────────────────────────────────────────────

uploaded = st.file_uploader("Upload IKEA Planner PDF", type=["pdf"])

if uploaded:
    with st.spinner("Parsing PDF..."):
        items = parse_ikea_pdf(uploaded)

    if not items:
        st.error("Could not parse any items. Make sure this is an IKEA Planner PDF.")
    else:
        st.success(f"✅ Found **{len(items)} items** across **{len(set(i['Cabinet / Category'] for i in items))} cabinets/categories**")

        import pandas as pd
        st.subheader("Select items for your quotation")
        st.markdown("Tick the items you want — they'll appear in Sheet 2 of the Excel.")

        df = pd.DataFrame(items)
        df.insert(0, 'Select', False)

        edited = st.data_editor(
            df[['Select','Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)']],
            column_config={"Select": st.column_config.CheckboxColumn("Select", default=False)},
            use_container_width=True,
            height=400,
            hide_index=True
        )

        selected_indices = edited.index[edited['Select'] == True].tolist()
        selected_count = len(selected_indices)
        selected_total = sum(items[i]['Total Price (EUR)'] for i in selected_indices)
        grand_total = sum(i['Total Price (EUR)'] for i in items)

        col1, col2 = st.columns(2)
        col1.metric("Grand Total", f"€{grand_total:,.2f}")
        col2.metric("Selected Total", f"€{selected_total:,.2f}", delta=f"{selected_count} items selected")

        for i in selected_indices:
            items[i]['Select (Y/N)'] = 'Y'

        excel_buf = build_excel(items, selected_indices)
        filename = uploaded.name.replace(".pdf", "_ItemList.xlsx")
        st.download_button(
            label="📥 Download Excel",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.markdown("---")
st.caption("Built for IKEA Planner PDFs · Supports METOD kitchen range")

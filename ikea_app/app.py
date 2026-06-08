import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="IKEA PDF → Excel", page_icon="🛋️", layout="centered")

st.title("🛋️ IKEA Planner PDF → Excel")
st.markdown("Upload your IKEA planner PDF and get a clean Excel file instantly.")

# ── Parser ────────────────────────────────────────────────────────────────────

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

                # Cabinet heading: "1. ME/MA 144 - ..."
                if re.match(r'^\d+\.\s+[A-Z]', line) and len(line) > 5:
                    current_cabinet = line
                    i += 1
                    continue

                # Item line: "PRODUCTNAME  Nx  €XX.XX"
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


# ── Excel Builder ─────────────────────────────────────────────────────────────

def build_excel(items):
    wb = Workbook()

    hdr_fill  = PatternFill("solid", fgColor="003087")
    alt_fill  = PatternFill("solid", fgColor="EEF3FF")
    tot_fill  = PatternFill("solid", fgColor="001F5E")
    note_fill = PatternFill("solid", fgColor="FFF9E6")
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

    # ── Sheet 1 ──────────────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "All Items"

    cols   = ['Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)','Select (Y/N)']
    widths = [42, 18, 50, 14, 6, 18, 18, 12]

    for ci, (h, w) in enumerate(zip(cols, widths), 1):
        hdr_cell(ws1, 1, ci, h, w)
    ws1.row_dimensions[1].height = 28
    ws1.freeze_panes = "A2"

    for ri, item in enumerate(items, 2):
        fill = alt_fill if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for ci, key in enumerate(cols, 1):
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

    # Grand total
    tr = len(items) + 2
    for ci in range(1, 9):
        c = ws1.cell(row=tr, column=ci)
        c.fill = tot_fill
        c.border = bdr
    ws1.cell(row=tr, column=6, value="GRAND TOTAL").font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    ws1.cell(row=tr, column=6).alignment = Alignment(horizontal="right")
    ws1.cell(row=tr, column=6).fill = tot_fill
    tc = ws1.cell(row=tr, column=7, value=f"=SUM(G2:G{tr-1})")
    tc.font = Font(name="Arial", bold=True, color="FFDD00", size=11)
    tc.number_format = '#,##0.00'
    tc.alignment = Alignment(horizontal="right")
    tc.fill = tot_fill

    # ── Sheet 2 ──────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Selected Items")

    note = ("HOW TO USE:  In the 'All Items' sheet, type  Y  in the 'Select (Y/N)' column "
            "for each item you want. Then copy those rows here into row 3 onwards.")
    ws2["A1"] = note
    ws2["A1"].font = Font(name="Arial", size=10, italic=True, color="555555")
    ws2["A1"].fill = note_fill
    ws2["A1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws2.merge_cells("A1:G1")
    ws2.row_dimensions[1].height = 40

    sel_cols   = ['Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)']
    sel_widths = [42, 18, 50, 14, 6, 18, 18]
    for ci, (h, w) in enumerate(zip(sel_cols, sel_widths), 1):
        hdr_cell(ws2, 2, ci, h, w)
    ws2.row_dimensions[2].height = 25

    # Paste area placeholder
    for ri in range(3, 8):
        for ci in range(1, 8):
            c = ws2.cell(row=ri, column=ci, value="← Paste selected rows here" if (ri == 3 and ci == 1) else "")
            c.font = Font(name="Arial", size=10, italic=True, color="AAAAAA")
            c.fill = PatternFill("solid", fgColor="FAFAFA")
            c.border = bdr

    # Selected total row
    for ci in range(1, 8):
        ws2.cell(row=9, column=ci).fill = tot_fill
        ws2.cell(row=9, column=ci).border = bdr
    ws2.cell(row=9, column=6, value="SELECTED TOTAL").font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    ws2.cell(row=9, column=6).alignment = Alignment(horizontal="right")
    ws2.cell(row=9, column=6).fill = tot_fill
    ws2.cell(row=9, column=7, value="=SUM(G3:G8)").font = Font(name="Arial", bold=True, color="FFDD00", size=11)
    ws2.cell(row=9, column=7).number_format = '#,##0.00'
    ws2.cell(row=9, column=7).alignment = Alignment(horizontal="right")
    ws2.cell(row=9, column=7).fill = tot_fill

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

        # Preview table
        st.subheader("Preview")
        import pandas as pd
        df = pd.DataFrame(items)[['Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)']]
        st.dataframe(df, use_container_width=True, height=350)

        total = sum(i['Total Price (EUR)'] for i in items)
        st.metric("Grand Total", f"€{total:,.2f}")

        # Download button
        excel_buf = build_excel(items)
        filename = uploaded.name.replace(".pdf", "_ItemList.xlsx")
        st.download_button(
            label="📥 Download Excel",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.info("💡 In the downloaded Excel, go to **All Items** sheet and type **Y** in the last column to flag items for your quotation, then copy them to **Selected Items** sheet.")

st.markdown("---")
st.caption("Built for IKEA Planner PDFs · Supports METOD kitchen range")

import streamlit as st
import pdfplumber
import re
import io
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="IKEA PDF → Excel", page_icon="🛋️", layout="wide")
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

            # Extract item images from this page
            small_imgs = sorted(
                [img for img in page.images if img['srcsize'][0] < 200],
                key=lambda x: x['top']
            )
            page_image_bufs = []
            for img_obj in small_imgs:
                try:
                    bbox = (img_obj['x0'], img_obj['top'], img_obj['x1'], img_obj['bottom'])
                    cropped = page.crop(bbox)
                    pil = cropped.to_image(resolution=96).original
                    buf = io.BytesIO()
                    pil.save(buf, format='PNG')
                    buf.seek(0)
                    page_image_bufs.append(buf)
                except:
                    page_image_bufs.append(None)

            lines = text.split('\n')
            img_index = 0
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

                    img_buf = page_image_bufs[img_index] if img_index < len(page_image_bufs) else None
                    img_index += 1

                    items.append({
                        'Cabinet / Category': current_cabinet,
                        'Product Name': prod_name,
                        'Description': ' '.join(desc_lines),
                        'Article No.': article_no or '',
                        'Qty': qty,
                        'Unit Price (EUR)': unit_price or round(total_price / qty, 2),
                        'Total Price (EUR)': total_price,
                        'Select (Y/N)': '',
                        '_image': img_buf
                    })
                    i = j
                    continue
                i += 1
    return items


def autofit_sheet(ws):
    """Autofit all columns based on content and center all cells."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                cell_len = len(str(cell.value)) if cell.value else 0
                if cell_len > max_len:
                    max_len = cell_len
            except:
                pass
        # Cap description col at 55, others reasonable max
        adjusted = min(max_len + 4, 55)
        adjusted = max(adjusted, 8)
        ws.column_dimensions[col_letter].width = adjusted


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

    def write_rows(ws, items_list, start_row=2, include_select=False):
        cols = ['Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)']
        if include_select:
            cols.append('Select (Y/N)')
        
        prev_cabinet = None
        for ri, item in enumerate(items_list, start_row):
            fill = alt_fill if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
            ws.row_dimensions[ri].height = 22

            # Row number column (col 1)
            c = ws.cell(row=ri, column=1, value=ri - start_row + 1)
            c.font = Font(name="Arial", size=10, bold=True, color="003087")
            c.fill = fill
            c.border = bdr
            c.alignment = Alignment(horizontal="center", vertical="center")

            # Data columns (col 2 onwards)
            for ci, key in enumerate(cols, 2):
                val = item.get(key, '')
                # Only show cabinet name on first row of each group
                if key == 'Cabinet / Category':
                    if val == prev_cabinet:
                        val = ''
                    else:
                        prev_cabinet = val
                c = ws.cell(row=ri, column=ci, value=val)
                c.font = Font(name="Arial", size=10, bold=(key == 'Cabinet / Category' and val != ''))
                c.fill = fill
                c.border = bdr
                if ci in (7, 8):
                    c.number_format = '#,##0.00'
                    c.alignment = Alignment(horizontal="right", vertical="center")
                elif key == 'Select (Y/N)':
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(vertical="center", wrap_text=False)

    n = len(items)

    # ── Sheet 1: All Items ────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "All Items"

    all_cols   = ['#', 'Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)','Select (Y/N)']

    for ci, h in enumerate(all_cols, 1):
        hdr_cell(ws1, 1, ci, h)
    ws1.row_dimensions[1].height = 32
    ws1.freeze_panes = "A2"

    # Y/N dropdown on col I
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True, showDropDown=False)
    dv.sqref = f"I2:I{n+1}"
    ws1.add_data_validation(dv)

    write_rows(ws1, items, start_row=2, include_select=True)

    # Grand total
    tr = n + 2
    for ci in range(1, 10):
        ws1.cell(row=tr, column=ci).fill = tot_fill
        ws1.cell(row=tr, column=ci).border = bdr
    ws1.cell(row=tr, column=7, value="GRAND TOTAL").font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    ws1.cell(row=tr, column=7).alignment = Alignment(horizontal="right")
    ws1.cell(row=tr, column=7).fill = tot_fill
    tc = ws1.cell(row=tr, column=8, value=f"=SUM(H2:H{tr-1})")
    tc.font = Font(name="Arial", bold=True, color="FFDD00", size=11)
    tc.number_format = '#,##0.00'
    tc.alignment = Alignment(horizontal="right")
    tc.fill = tot_fill

    autofit_sheet(ws1)

    # ── Sheet 2: Selected Items ───────────────────────────────────────────────
    ws2 = wb.create_sheet("Selected Items")

    sel_cols   = ['#', 'Cabinet / Category','Product Name','Description','Article No.','Qty','Unit Price (EUR)','Total Price (EUR)']
    for ci, h in enumerate(sel_cols, 1):
        hdr_cell(ws2, 1, ci, h)
    ws2.row_dimensions[1].height = 32

    selected = [items[i] for i in selected_indices] if selected_indices else []

    if selected:
        write_rows(ws2, selected, start_row=2, include_select=False)
        sel_total_row = len(selected) + 2
    else:
        ws2.cell(row=2, column=1, value="No items selected. Go to All Items sheet and type Y in column I.").font = Font(name="Arial", size=10, italic=True, color="999999")
        ws2.merge_cells("A2:H2")
        sel_total_row = 3

    for ci in range(1, 9):
        ws2.cell(row=sel_total_row, column=ci).fill = tot_fill
        ws2.cell(row=sel_total_row, column=ci).border = bdr
    ws2.cell(row=sel_total_row, column=7, value="SELECTED TOTAL").font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    ws2.cell(row=sel_total_row, column=7).alignment = Alignment(horizontal="right")
    ws2.cell(row=sel_total_row, column=7).fill = tot_fill
    sel_total = sum(items[i]['Total Price (EUR)'] for i in (selected_indices or []))
    tc2 = ws2.cell(row=sel_total_row, column=8, value=sel_total)
    tc2.font = Font(name="Arial", bold=True, color="FFDD00", size=11)
    tc2.number_format = '#,##0.00'
    tc2.alignment = Alignment(horizontal="right")
    tc2.fill = tot_fill

    autofit_sheet(ws2)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── UI ────────────────────────────────────────────────────────────────────────

uploaded = st.file_uploader("Upload IKEA Planner PDF", type=["pdf"])

if uploaded:
    with st.spinner("Parsing PDF and extracting images..."):
        items = parse_ikea_pdf(uploaded)

    if not items:
        st.error("Could not parse any items. Make sure this is an IKEA Planner PDF.")
    else:
        st.success(f"✅ Found **{len(items)} items** across **{len(set(i['Cabinet / Category'] for i in items))} categories**")

        import pandas as pd
        st.subheader("Select items for your quotation")

        # Init session state for persistent selections
        if 'selected_set' not in st.session_state:
            st.session_state.selected_set = set()

        # Category filter
        all_categories = sorted(set(i['Cabinet / Category'] for i in items), key=lambda x: int(x.split('.')[0]) if x.split('.')[0].isdigit() else 999)
        selected_cats = st.multiselect(
            "Filter by category (leave blank to show all)",
            options=all_categories,
            default=[]
        )
        filtered_items = items if not selected_cats else [i for i in items if i['Cabinet / Category'] in selected_cats]
        filtered_indices = [items.index(i) for i in filtered_items]

        display_df = pd.DataFrame([{
            'Select': idx in st.session_state.selected_set,
            'Product Name': i['Product Name'],
            'Description': i['Description'],
            'Article No.': i['Article No.'],
            'Qty': i['Qty'],
            'Unit Price': i['Unit Price (EUR)'],
            'Total': i['Total Price (EUR)'],
        } for idx, i in zip(filtered_indices, filtered_items)])

        editor_key = "editor_" + "_".join(selected_cats)

        def sync_selections():
            edited_state = st.session_state[editor_key]
            if "edited_rows" in edited_state:
                for row_str, changes in edited_state["edited_rows"].items():
                    row_i = int(row_str)
                    real_idx = filtered_indices[row_i]
                    if "Select" in changes:
                        if changes["Select"]:
                            st.session_state.selected_set.add(real_idx)
                        else:
                            st.session_state.selected_set.discard(real_idx)

        st.data_editor(
            display_df,
            column_config={
                "Select": st.column_config.CheckboxColumn("✅ Select", default=False, width="small"),
                "Product Name": st.column_config.TextColumn("Product Name", width="medium"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Article No.": st.column_config.TextColumn("Article No.", width="small"),
                "Qty": st.column_config.NumberColumn("Qty", width="small"),
                "Unit Price": st.column_config.NumberColumn("Unit Price (€)", width="small", format="€%.2f"),
                "Total": st.column_config.NumberColumn("Total (€)", width="small", format="€%.2f"),
            },
            use_container_width=True,
            height=500,
            hide_index=True,
            key=editor_key,
            on_change=sync_selections
        )

        selected_indices = list(st.session_state.selected_set)
        selected_total = sum(items[i]['Total Price (EUR)'] for i in selected_indices)
        grand_total = sum(i['Total Price (EUR)'] for i in items)

        col1, col2 = st.columns(2)
        col1.metric("Grand Total", f"€{grand_total:,.2f}")
        col2.metric("Selected Total", f"€{selected_total:,.2f}", delta=f"{len(selected_indices)} items selected")

        for i in selected_indices:
            items[i]['Select (Y/N)'] = 'Y'

        if st.button("📥 Generate & Download Excel", use_container_width=True, type="primary"):
            with st.spinner("Building Excel with images..."):
                excel_buf = build_excel(items, selected_indices)
            filename = uploaded.name.replace(".pdf", "_ItemList.xlsx")
            st.download_button(
                label="⬇️ Click to Download",
                data=excel_buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

st.markdown("---")
st.caption("Built for IKEA Planner PDFs · Supports METOD kitchen range")

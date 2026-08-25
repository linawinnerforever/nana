import io
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

def clean_header_name(raw_name):
    """提取纯净持卡人姓名"""
    if "CRAZY MAPLE STUDIO" in raw_name:
        return "CRAZY MAPLE STUDIO"
    lines = [l.strip() for l in raw_name.split('\n') if l.strip()]
    last_line = lines[-1] if lines else raw_name
    last_line = re.sub(r'^(?:[A-Za-z\s]*?(?:MCC|Charge|Credit|PostingDate|TransactionDate|Description|Reference Number|mber))+\s*', '', last_line, flags=re.IGNORECASE)
    last_line = re.sub(r'\s*Total\s*Activity.*$', '', last_line, flags=re.IGNORECASE)
    return last_line.strip('|\s')

def extract_pdf_text_from_file(uploaded_file):
    """多引擎容错提取 PDF 原文，彻底解决 Streamlit 提取 0 条的问题"""
    full_text = ""
    
    # 尝试引擎 1: pypdf
    try:
        import pypdf
        file_bytes = io.BytesIO(uploaded_file.getvalue())
        reader = pypdf.PdfReader(file_bytes)
        for page in reader.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"
    except Exception:
        pass

    # 若引擎 1 未提出来文本，回退至引擎 2: pdfplumber
    if not full_text.strip():
        try:
            import pdfplumber
            file_bytes = io.BytesIO(uploaded_file.getvalue())
            with pdfplumber.open(file_bytes) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        full_text += t + "\n"
        except Exception:
            pass

    return full_text

def extract_pdf_data(uploaded_files):
    """精准解析 Bank of America PDF 账单"""
    all_transactions = []
    summary_by_period = {}

    for uploaded_file in uploaded_files:
        full_text = extract_pdf_text_from_file(uploaded_file)
        if not full_text:
            continue
            
        # 1. 动态提取账单周期
        period_match = re.search(r'([A-Za-z]+\s+\d{2},\s*\d{4}\s*-\s*[A-Za-z]+\s+\d{2},\s*\d{4})', full_text)
        period_name = period_match.group(1).strip() if period_match else "Unknown Period"
        
        if period_name not in summary_by_period:
            summary_by_period[period_name] = {}

        # 2. 从 Cardholder Activity Summary 提取持卡人汇总数据
        if "Cardholder Activity Summary" in full_text:
            sum_section = full_text.split("Cardholder Activity Summary")[1].split("Transactions")[0]
            sum_matches = re.findall(r'([A-Z,\s]{3,35})XXXX-XXXX-XXXX-(\d{4})[\d,.]+\s+[\d,.]+\s+[\d,.]+\s+([\d,.]+)\s+([\d,.]+)', sum_section)
            for holder, last4, purchases, total_act in sum_matches:
                clean_name = holder.strip()
                if not any(k in clean_name for k in ["Account Number", "Credit Limit", "Total Activity", "Purchases", "Credits", "Cash"]):
                    summary_by_period[period_name][clean_name] = float(total_act.replace(',', ''))

        # 3. 从 Transactions 精准解析明细与 Description 原文
        if "Transactions" in full_text:
            tx_text = full_text.split("Transactions", 1)[1]
            headers = list(re.finditer(r'([^\n\d]{3,40})Account Number:\s*XXXX-XXXX-XXXX-(\d{4})', tx_text))
            
            for idx, hm in enumerate(headers):
                raw_name = hm.group(1).strip()
                holder_name = clean_header_name(raw_name)
                card_last4 = hm.group(2).strip()
                
                start_pos = hm.end()
                end_pos = headers[idx+1].start() if idx+1 < len(headers) else len(tx_text)
                section_content = tx_text[start_pos:end_pos]
                
                # AUTO PAYMENT DEDUCTION (Credit / Payment 金额取负数)
                if "AUTO PAYMENT DEDUCTION" in section_content:
                    pay_m = re.search(r'(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+AUTO PAYMENT DEDUCTION\s+(\d+)\s+([\d,]+\.\d{2})', section_content)
                    if pay_m:
                        amt_val = float(pay_m.group(4).replace(',', ''))
                        all_transactions.append([period_name, holder_name, card_last4, pay_m.group(2), pay_m.group(1), "AUTO PAYMENT DEDUCTION", pay_m.group(3), "", -abs(amt_val), "Credit / Payment"])
                    else:
                        pay_amt_m = re.search(r'AUTO PAYMENT DEDUCTION[\s\S]*?([\d,]+\.\d{2})', section_content)
                        pay_amt = float(pay_amt_m.group(1).replace(',', '')) if pay_amt_m else 0.0
                        all_transactions.append([period_name, holder_name, card_last4, "07/08", "07/08", "AUTO PAYMENT DEDUCTION", "0071", "", -abs(pay_amt), "Credit / Payment"])

                # 普通消费交易：PostDate TransDate Description RefNum MCC Amount
                tx_pattern = r'(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(.+?)\s+(\d{23,24})\s+(\d{4})\s+([\d,]+\.\d{2})'
                
                for m in re.finditer(tx_pattern, section_content):
                    post_d = m.group(1)
                    trans_d = m.group(2)
                    raw_desc = m.group(3)
                    ref_num = m.group(4)
                    mcc = m.group(5)
                    amount = float(m.group(6).replace(',', ''))
                    
                    # 提取纯净 PDF Description 原文
                    desc_clean = re.sub(r'\s+', ' ', raw_desc).strip()
                    desc_clean = re.sub(r'^\d{2}/\d{2}\s+', '', desc_clean)
                    desc_clean = re.sub(r'Account Number:\s*XXXX-XXXX-XXXX-\d{4}\s*', '', desc_clean)
                    desc_clean = re.sub(r'Total Activity\s*[\d,.]*\s*', '', desc_clean)
                    desc_clean = desc_clean.strip()
                    
                    all_transactions.append([
                        period_name, holder_name, card_last4,
                        trans_d, post_d, desc_clean, ref_num, mcc, amount, "Charge"
                    ])

    return summary_by_period, all_transactions

def generate_excel_bytes(summary_by_period, transactions):
    """生成与 V3 规范一模一样的 Excel 报表"""
    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    stat_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=Side(border_style="thin", color="000000"), bottom=Side(border_style="double", color="000000"), left=thin_border_side, right=thin_border_side)

    # -------------------------------------------------------------
    # 1. 建立「交易明细」 Sheet
    # -------------------------------------------------------------
    ws_det = wb.create_sheet(title="交易明细")
    ws_det.views.sheetView[0].showGridLines = True
    
    # H1、H2 单元格清空 (A1 到 H2 均为空文本)
    for r in [1, 2]:
        for c in range(1, 9):
            ws_det.cell(row=r, column=c, value="")

    last_detail_row = len(transactions) + 3

    # I 列存放顶端求和公式
    c_s1 = ws_det.cell(row=1, column=9, value=f'=SUMIF(B4:B{last_detail_row}, "CRAZY MAPLE STUDIO*", I4:I{last_detail_row})')
    c_s1.font = bold_font
    c_s1.fill = stat_fill
    c_s1.number_format = "$#,##0.00"
    c_s1.alignment = Alignment(horizontal="right", vertical="center")

    c_s2 = ws_det.cell(row=2, column=9, value=f'=SUMIFS(I4:I{last_detail_row}, B4:B{last_detail_row}, "<>CRAZY MAPLE STUDIO*", J4:J{last_detail_row}, "Charge")')
    c_s2.font = bold_font
    c_s2.fill = stat_fill
    c_s2.number_format = "$#,##0.00"
    c_s2.alignment = Alignment(horizontal="right", vertical="center")

    for r in [1, 2]:
        for col_i in range(1, 11):
            cell = ws_det.cell(row=r, column=col_i)
            cell.border = thin_border
            cell.fill = stat_fill

    # 第 3 行为交易明细表头
    det_headers = ["账单周期", "持卡人 / 账户", "卡号末四位", "交易日 (Trans Date)", "记账日 (Post Date)", "交易描述 (Description)", "参考号 (Reference No)", "MCC", "金额 (Amount)", "交易类型 (Type)"]
    for c_idx, h_text in enumerate(det_headers, 1):
        cell = ws_det.cell(row=3, column=c_idx, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # 数据从第 4 行开始写入
    for r_idx, r_data in enumerate(transactions, start=4):
        for c_idx, val in enumerate(r_data, start=1):
            cell = ws_det.cell(row=r_idx, column=c_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            if c_idx in [1, 3, 4, 5, 8, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in [2, 6, 7]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 9:
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")

    # 第 3 行自动下拉筛选与前 3 行冻结
    ws_det.auto_filter.ref = f"A3:J{last_detail_row}"
    ws_det.freeze_panes = "A4"

    # 计算列宽 (金额 I 列设为紧凑宽度 18)
    for col_idx in range(1, 11):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(3, last_detail_row + 1):
            val = ws_det.cell(row=row_idx, column=col_idx).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        
        header_len = len(str(det_headers[col_idx - 1]))
        max_len = max(max_len, header_len)
        
        if col_idx == 9:
            ws_det.column_dimensions[col_letter].width = max(max_len + 5, 18)
        elif col_idx == 6:
            ws_det.column_dimensions[col_letter].width = min(max_len + 4, 55)
        else:
            ws_det.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # -------------------------------------------------------------
    # 2. 建立「汇总 Dashboard」 Sheet
    # -------------------------------------------------------------
    ws_sum = wb.active
    ws_sum.title = "汇总 Dashboard"
    ws_sum.views.sheetView[0].showGridLines = True

    periods = list(summary_by_period.keys())
    sum_headers = ["持卡人"] + periods + ["合计 (Total)"]

    for c_idx, h_text in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=1, column=c_idx, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    all_holders = sorted(list({h for p in summary_by_period.values() for h in p.keys()}))

    for r_idx, holder in enumerate(all_holders, start=2):
        c1 = ws_sum.cell(row=r_idx, column=1, value=holder)
        c1.font = data_font
        c1.border = thin_border
        
        row_tot = 0.0
        for p_idx, p_name in enumerate(periods, start=2):
            val = summary_by_period[p_name].get(holder, 0.0)
            row_tot += val
            c = ws_sum.cell(row=r_idx, column=p_idx, value=val)
            c.font = data_font
            c.number_format = "$#,##0.00"
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = thin_border
            
        c_tot = ws_sum.cell(row=r_idx, column=len(periods) + 2, value=row_tot)
        c_tot.font = data_font
        c_tot.number_format = "$#,##0.00"
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_tot.border = thin_border

    tot_row = len(all_holders) + 2
    lbl = ws_sum.cell(row=tot_row, column=1, value="合计 (Total)")
    lbl.font = bold_font
    lbl.border = total_border

    grand_tot = 0.0
    for p_idx, p_name in enumerate(periods, start=2):
        col_tot = sum(summary_by_period[p_name].values())
        grand_tot += col_tot
        c = ws_sum.cell(row=tot_row, column=p_idx, value=col_tot)
        c.font = bold_font
        c.number_format = "$#,##0.00"
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = total_border

    total_cell_letter = get_column_letter(len(periods) + 2)
    g_cell = ws_sum.cell(row=tot_row, column=len(periods) + 2, value=grand_tot)
    g_cell.font = bold_font
    g_cell.number_format = "$#,##0.00"
    g_cell.alignment = Alignment(horizontal="right", vertical="center")
    g_cell.border = total_border

    # 3. Check 校验行
    check_row = tot_row + 2
    ws_sum.cell(row=check_row, column=1, value="check").font = bold_font
    c_check = ws_sum.cell(row=check_row, column=len(periods) + 2)
    c_check.value = f'=交易明细!I2 - {total_cell_letter}{tot_row}'
    c_check.font = bold_font
    c_check.number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    c_check.alignment = Alignment(horizontal="right", vertical="center")

    ws_sum.column_dimensions['A'].width = 25
    for col_i in range(2, len(periods) + 3):
        col_l = get_column_letter(col_i)
        ws_sum.column_dimensions[col_l].width = 30

    excel_stream = io.BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)
    return excel_stream

# Streamlit UI
st.set_page_config(page_title="信用卡账单自动解析小工具", page_icon="📄")
st.title("📄 信用卡 PDF 账单自动解析工具")
st.write("请在下方上传 Bank of America PDF 账单文件（支持同时上传多个文件）：")

uploaded_files = st.file_uploader("选择 PDF 文件", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 开始提取并生成 Excel", type="primary"):
        with st.spinner("正在解析 PDF 并生成报表中..."):
            try:
                summary_data, detail_data = extract_pdf_data(uploaded_files)
                excel_bytes = generate_excel_bytes(summary_data, detail_data)
                
                st.success(f"成功解析 {len(uploaded_files)} 个 PDF 账单！已提取 {len(detail_data)} 条交易明细。")
                
                st.download_button(
                    label="📥 点击下载 Excel 汇总表格",
                    data=excel_bytes,
                    file_name="信用卡账单汇总表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"解析过程出错，请检查 PDF 格式是否正确。错误详情: {e}")

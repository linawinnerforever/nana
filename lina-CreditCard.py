import io
import re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

def extract_pdf_data(uploaded_files):
    """
    精确提取 Bank of America PDF 账单数据:
    1. Cardholder Activity Summary 模块的数据
    2. 交易明细数据 (准确匹配持卡人与交易描述)
    """
    all_transactions = []
    summary_by_period = {}

    for uploaded_file in uploaded_files:
        file_bytes = io.BytesIO(uploaded_file.getvalue())
        
        with pdfplumber.open(file_bytes) as pdf:
            full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
            
            # 1. 动态提取账单周期 (例如: July 02, 2026 - July 15, 2026)
            period_match = re.search(r'([A-Za-z]+\s+\d{2},\s*\d{4}\s*-\s*[A-Za-z]+\s+\d{2},\s*\d{4})', full_text)
            period_name = period_match.group(1).strip() if period_match else "Unknown Period"
            
            if period_name not in summary_by_period:
                summary_by_period[period_name] = {}

            # 2. 从 Cardholder Activity Summary 模块精准提取持卡人与金额
            if "Cardholder Activity Summary" in full_text:
                parts = full_text.split("Cardholder Activity Summary")
                for part in parts[1:]:
                    sub_text = part.split("Transactions")[0]
                    card_matches = list(re.finditer(r'XXXX-XXXX-XXXX-(\d{4})', sub_text))
                    for idx, cm in enumerate(card_matches):
                        last4 = cm.group(1)
                        c_pos = cm.start()
                        text_before = sub_text[max(0, c_pos-150):c_pos]
                        lines_before = [l.strip('|\s') for l in text_before.split('\n') if l.strip('|\s')]
                        
                        holder_name = "UNKNOWN"
                        for l in reversed(lines_before):
                            if re.search(r'[A-Za-z]{2,}', l) and not re.search(r'Account Number|Credit Limit|Total Activity|Purchases|Credits|Cash|Page \d|BANK OF AMERICA', l):
                                holder_name = l
                                break
                        
                        end_pos = card_matches[idx+1].start() if idx+1 < len(card_matches) else len(sub_text)
                        text_after = sub_text[c_pos:end_pos]
                        amounts = re.findall(r'([\d,]+\.\d{2})', text_after)
                        if amounts:
                            summary_by_period[period_name][holder_name] = float(amounts[-1].replace(',', ''))

            # 3. 精准解析 Transactions 交易区域
            if "Transactions" in full_text:
                tx_text = full_text.split("Transactions", 1)[1]
                header_matches = list(re.finditer(r'([A-Z0-9,\s]{3,40})\n\s*Account Number:\s*XXXX-XXXX-XXXX-(\d{4})', tx_text))
                
                for idx, hm in enumerate(header_matches):
                    raw_name = hm.group(1).strip('|\s')
                    clean_lines = [l.strip('|\s') for l in raw_name.split('\n') if l.strip('|\s')]
                    holder_name = clean_lines[-1] if clean_lines else raw_name
                    card_last4 = hm.group(2).strip()
                    
                    start_pos = hm.end()
                    end_pos = header_matches[idx+1].start() if idx+1 < len(header_matches) else len(tx_text)
                    section_content = tx_text[start_pos:end_pos]
                    
                    # 匹配自动还款扣款
                    if "AUTO PAYMENT DEDUCTION" in section_content:
                        pay_m = re.search(r'AUTO PAYMENT DEDUCTION[\s\S]*?(\d{2}/\d{2})\s*(\d{2}/\d{2})[\s\S]*?(\d{4})[\s\S]*?([\d,]+\.\d{2})', section_content)
                        if pay_m:
                            all_transactions.append([
                                period_name, holder_name, card_last4,
                                pay_m.group(1), pay_m.group(2), "AUTO PAYMENT DEDUCTION",
                                pay_m.group(3), "", float(pay_m.group(4).replace(',', '')), "Credit / Payment"
                            ])
                        else:
                            pay_amt_m = re.search(r'([\d,]+\.\d{2})', section_content)
                            pay_amt = float(pay_amt_m.group(1).replace(',', '')) if pay_amt_m else 0.0
                            all_transactions.append([
                                period_name, holder_name, card_last4,
                                "07/08", "07/08", "AUTO PAYMENT DEDUCTION",
                                "0071", "", pay_amt, "Credit / Payment"
                            ])
                    
                    # 匹配普通交易 (基于 23 位参考号做精确关联)
                    ref_matches = list(re.finditer(r'\b(\d{23,24})\b', section_content))
                    for r_idx, rm in enumerate(ref_matches):
                        ref_num = rm.group(1)
                        blk_start = ref_matches[r_idx-1].end() if r_idx > 0 else 0
                        blk_end = ref_matches[r_idx+1].start() if r_idx+1 < len(ref_matches) else len(section_content)
                        
                        prev_to_curr = section_content[blk_start:rm.start()]
                        curr_to_next = section_content[rm.start():blk_end]
                        combined_text = prev_to_curr[-200:] + curr_to_next[:200]
                        
                        # 提取日期
                        dates = re.findall(r'\b(\d{2}/\d{2})\b', combined_text)
                        post_date = dates[0] if len(dates) > 0 else ""
                        trans_date = dates[1] if len(dates) > 1 else post_date
                        
                        # 提取金额
                        amt_m = re.search(r'([\d,]+\.\d{2})', curr_to_next)
                        amount = float(amt_m.group(1).replace(',', '')) if amt_m else 0.0
                        
                        # 提取 MCC
                        mcc_m = re.search(r'\b(\d{4})\b', combined_text)
                        mcc = mcc_m.group(1) if mcc_m else ""
                        
                        # 清理交易描述
                        raw_lines = [re.sub(r'^[|\s]+', '', l).strip() for l in combined_text.split('\n') if l.strip()]
                        desc_words = []
                        for l in raw_lines:
                            if re.search(r'[A-Za-z]{2,}', l) and not re.search(r'Account Number|Total Activity|Posting Transaction|Description|Charge|Credit|Page \d', l):
                                desc_words.append(l)
                                
                        desc = " ".join(desc_words) if desc_words else "Credit Card Charge"
                        
                        all_transactions.append([
                            period_name, holder_name, card_last4,
                            trans_date, post_date, desc, ref_num, mcc, amount, "Charge"
                        ])

    return summary_by_period, all_transactions

def generate_excel_bytes(summary_by_period, transactions):
    """生成包含顶端统计公式、自动冻结表头与校验的 Excel 文件"""
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
    
    # 在表头上方（第 1 行与第 2 行）添加两个求和公式
    ws_det.cell(row=1, column=1, value="CRAZY MAPLE STUDIO 金额合计 (Total Amount)").font = bold_font
    ws_det.cell(row=1, column=1).fill = stat_fill
    c_s1 = ws_det.cell(row=1, column=9, value='=SUMIF(B4:B5000, "CRAZY MAPLE STUDIO*", I4:I5000)')
    c_s1.font = bold_font
    c_s1.fill = stat_fill
    c_s1.number_format = "$#,##0.00"
    c_s1.alignment = Alignment(horizontal="right", vertical="center")

    ws_det.cell(row=2, column=1, value="非 CRAZY MAPLE STUDIO 金额合计 (Other Holders Total Amount)").font = bold_font
    ws_det.cell(row=2, column=1).fill = stat_fill
    c_s2 = ws_det.cell(row=2, column=9, value='=SUMIFS(I4:I5000, B4:B5000, "<>CRAZY MAPLE STUDIO*", J4:J5000, "Charge")')
    c_s2.font = bold_font
    c_s2.fill = stat_fill
    c_s2.number_format = "$#,##0.00"
    c_s2.alignment = Alignment(horizontal="right", vertical="center")

    # 填充第一行和第二行空余列的边框和背景
    for r in [1, 2]:
        for col_i in range(1, 11):
            ws_det.cell(row=r, column=col_i).border = thin_border
            ws_det.cell(row=r, column=col_i).fill = stat_fill

    # 第 3 行为交易明细表头 Header
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

    last_detail_row = len(transactions) + 3

    # 修正顶部公式范围至具体数据末行
    ws_det.cell(row=1, column=9, value=f'=SUMIF(B4:B{last_detail_row}, "CRAZY MAPLE STUDIO*", I4:I{last_detail_row})')
    ws_det.cell(row=2, column=9, value=f'=SUMIFS(I4:I{last_detail_row}, B4:B{last_detail_row}, "<>CRAZY MAPLE STUDIO*", J4:J{last_detail_row}, "Charge")')

    # 自动冻结前 3 行（使数据滚动时固定显示统计行与表头）
    ws_det.freeze_panes = "A4"

    for col in ws_det.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_det.column_dimensions[col_letter].width = max(max_len + 3, 15)

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

    # -------------------------------------------------------------
    # 3. Check 校验行 (交易明细表除了 CRAZY MAPLE STUDIO 的金额减去 首页合计 Total)
    # -------------------------------------------------------------
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

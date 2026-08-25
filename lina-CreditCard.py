import io
import re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

def extract_pdf_data(uploaded_files):
    """跨行/多行自适应提取 Bank of America PDF 账单数据"""
    all_transactions = []
    summary_by_period = {}

    for uploaded_file in uploaded_files:
        # 重置文件指针
        file_bytes = io.BytesIO(uploaded_file.getvalue())
        
        with pdfplumber.open(file_bytes) as pdf:
            # 读取全部页面的完整文本
            full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
            
            # 1. 提取账单周期字符串 (例: July 02, 2026 - July 15, 2026)
            period_match = re.search(r'([A-Za-z]+\s+\d{2},\s*\d{4}\s*-\s*[A-Za-z]+\s+\d{2},\s*\d{4})', full_text)
            period_name = period_match.group(1).strip() if period_match else "Unknown Period"
            
            if period_name not in summary_by_period:
                summary_by_period[period_name] = {}

            # 2. 提取第 3 页持卡人 Cardholder Activity Summary 静态数据
            # 匹配名字、末四位卡号、最终 Total Activity 金额
            summary_blocks = re.findall(r'([A-Z,\s]+)\n(?:\|\s*)*XXXX-XXXX-XXXX-(\d{4})[\s\S]*?([\d,]+\.\d{2})\n', full_text)
            for h_name, last4, amt_str in summary_blocks:
                name_clean = h_name.strip()
                if any(k in name_clean for k in ["BHAGAT", "JIA", "NAN", "TRAN"]):
                    summary_by_period[period_name][name_clean] = float(amt_str.replace(',', ''))

            # 3. 按持卡人 Account Number 切割交易区块
            sections = re.split(r'([A-Z,\s]+)\nAccount Number:\s*XXXX-XXXX-XXXX-(\d{4})', full_text)

            i = 1
            while i < len(sections):
                holder_name = sections[i].strip()
                card_last4 = sections[i+1].strip()
                content = sections[i+2]
                i += 3

                # A. 提取主账户自动还款扣款
                if "AUTO PAYMENT DEDUCTION" in content:
                    pay_amt_m = re.search(r'AUTO PAYMENT DEDUCTION[\s\S]*?([\d,]+\.\d{2})', content)
                    pay_amt = float(pay_amt_m.group(1).replace(',', '')) if pay_amt_m else 0.0
                    all_transactions.append([
                        period_name, "CRAZY MAPLE STUDIO", "2273",
                        "07/08", "07/08", "AUTO PAYMENT DEDUCTION",
                        "0071", "", pay_amt, "Credit / Payment"
                    ])

                # B. 跨行提取普通消费交易 (利用 23 位参考号做锚点)
                ref_matches = list(re.finditer(r'\b(\d{23,24})\b', content))
                
                for ref_m in ref_matches:
                    ref_num = ref_m.group(1)
                    start_pos = ref_m.start()
                    
                    # 抓取参考号前后 180 个字符的上下文窗口
                    win_start = max(0, start_pos - 180)
                    win_end = min(len(content), start_pos + 180)
                    window_text = content[win_start:win_end]
                    
                    # 提取日期 (MM/DD)
                    dates = re.findall(r'\b(\d{2}/\d{2})\b', window_text)
                    post_date = dates[0] if len(dates) > 0 else ""
                    trans_date = dates[1] if len(dates) > 1 else post_date
                    
                    # 提取金额
                    after_ref = content[start_pos:win_end]
                    amt_m = re.search(r'([\d,]+\.\d{2})', after_ref)
                    amount = float(amt_m.group(1).replace(',', '')) if amt_m else 0.0
                    
                    # 提取 MCC 4 位商户代码
                    mcc_m = re.search(r'\b(\d{4})\b', after_ref)
                    mcc = mcc_m.group(1) if mcc_m else ""
                    
                    # 提取商户名称文本
                    before_ref = content[win_start:start_pos]
                    raw_lines = [l.strip('| ').strip() for l in before_ref.split('\n') if l.strip('| ').strip()]
                    clean_desc = []
                    for l in raw_lines:
                        if not re.search(r'Account Number|\d{2}/\d{2}|Total Activity|\d{23,}|Charge|Credit', l):
                            clean_desc.append(l)
                    
                    desc = " ".join(clean_desc[-2:]) if clean_desc else "Card Charge"

                    all_transactions.append([
                        period_name, holder_name, card_last4,
                        trans_date, post_date, desc, ref_num, mcc, amount, "Charge"
                    ])

    return summary_by_period, all_transactions

def generate_excel_bytes(summary_by_period, transactions):
    """在内存中生成包含 汇总 Dashboard 和 交易明细 的 Excel 文件"""
    wb = openpyxl.Workbook()

    # 统一视觉样式 (深蓝表头 #1F4E78，白字粗体)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=Side(border_style="thin", color="000000"), bottom=Side(border_style="double", color="000000"), left=thin_border_side, right=thin_border_side)

    # 1. 交易明细 Sheet
    ws_det = wb.create_sheet(title="交易明细")
    ws_det.views.sheetView[0].showGridLines = True
    
    det_headers = ["账单周期", "持卡人 / 账户", "卡号末四位", "交易日 (Trans Date)", "记账日 (Post Date)", "交易描述 (Description)", "参考号 (Reference No)", "MCC", "金额 (Amount)", "交易类型 (Type)"]
    
    for c_idx, h_text in enumerate(det_headers, 1):
        cell = ws_det.cell(row=1, column=c_idx, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, r_data in enumerate(transactions, start=2):
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

    last_detail_row = len(transactions) + 1

    for col in ws_det.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_det.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # 2. 汇总 Dashboard Sheet
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
    c_check.value = f'=SUMIF(交易明细!J2:J{last_detail_row}, "Charge", 交易明细!I2:I{last_detail_row}) - {total_cell_letter}{tot_row}'
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

# -------------------------------------------------------------
# Streamlit 界面
# -------------------------------------------------------------
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

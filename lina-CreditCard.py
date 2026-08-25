import re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def extract_pdf_data(pdf_paths):
    """
    从给定的 Bank of America 信用卡 PDF 账单中提取:
    1. 动态账单周期 (Billing Period)
    2. 页码 3 中的 Cardholder Activity Summary 静态汇总数据
    3. 详细交易明细 (Transactions)
    """
    all_transactions = []
    summary_by_period = {}  # 格式: {period_name: {cardholder_name: total_activity}}

    for pdf_path in pdf_paths:
        with pdfplumber.open(pdf_path) as pdf:
            period_name = "Unknown Period"
            
            # 1. 动态提取账单周期字符串 (例: July 02, 2026 - July 15, 2026)
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                period_match = re.search(r'([A-Za-z]+\s+\d{2},\s*\d{4}\s*-\s*[A-Za-z]+\s+\d{2},\s*\d{4})', text)
                if period_match:
                    period_name = period_match.group(1).strip()
                    break
            
            if period_name not in summary_by_period:
                summary_by_period[period_name] = {}

            current_cardholder = None
            current_card_last4 = None

            # 2. 逐页读取明细与汇总数据
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    line = line.strip()
                    
                    # 匹配持卡人 Header (例如: BHAGAT, KRUTI V Account Number: XXXX-XXXX-XXXX-4574)
                    card_match = re.search(r'([A-Z,\s]+)\s+Account Number:\s*XXXX-XXXX-XXXX-(\d{4})', line)
                    if card_match:
                        current_cardholder = card_match.group(1).strip()
                        current_card_last4 = card_match.group(2).strip()
                        continue
                    
                    # 提取 PDF Summary 区域静态数值 (Cardholder Activity Summary)
                    summary_match = re.search(r'^([A-Z,\s]+)\s+XXXX-XXXX-XXXX-\d{4}\s+[\d,.]+\s+[\d,.]+\s+[\d,.]+\s+([\d,.]+)\s+([\d,.]+)$', line)
                    if summary_match:
                        holder_name = summary_match.group(1).strip()
                        total_act = float(summary_match.group(3).replace(',', ''))
                        summary_by_period[period_name][holder_name] = total_act

                    # 提取普通交易明细行 (Date PostDate Description RefNum MCC Amount)
                    tx_match = re.search(r'^(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(.+?)\s+(\d{10,})\s+(\d{4})\s+([\d,]+\.\d{2})$', line)
                    if tx_match and current_cardholder:
                        all_transactions.append([
                            period_name, current_cardholder, current_card_last4,
                            tx_match.group(1), tx_match.group(2), tx_match.group(3),
                            tx_match.group(4), tx_match.group(5),
                            float(tx_match.group(6).replace(',', '')), "Charge"
                        ])

                    # 提取主账户自动还款扣款行 (AUTO PAYMENT DEDUCTION)
                    pay_match = re.search(r'AUTO PAYMENT DEDUCTION\s+(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(\d+)\s+([\d,]+\.\d{2})', line)
                    if pay_match:
                        all_transactions.append([
                            period_name, "CRAZY MAPLE STUDIO", "2273",
                            pay_match.group(1), pay_match.group(2), "AUTO PAYMENT DEDUCTION",
                            pay_match.group(3), "", float(pay_match.group(4).replace(',', '')), "Credit / Payment"
                        ])

    return summary_by_period, all_transactions

def generate_excel_report(summary_by_period, transactions, output_filename="bank_statement_final.xlsx"):
    wb = openpyxl.Workbook()

    # 统一视觉样式定义 (与明细表完全一致的深蓝表头)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=Side(border_style="thin", color="000000"), bottom=Side(border_style="double", color="000000"), left=thin_border_side, right=thin_border_side)

    # -------------------------------------------------------------
    # 1. 建立「交易明细」 Sheet (提前写入以确定 Check 公式的计算总行数)
    # -------------------------------------------------------------
    ws_det = wb.create_sheet(title="交易明细")
    ws_det.views.sheetView[0].showGridLines = True
    
    det_headers = [
        "账单周期", "持卡人 / 账户", "卡号末四位", "交易日 (Trans Date)", 
        "记账日 (Post Date)", "交易描述 (Description)", "参考号 (Reference No)", 
        "MCC", "金额 (Amount)", "交易类型 (Type)"
    ]
    
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

    # 自动调整明细列宽
    for col in ws_det.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_det.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # -------------------------------------------------------------
    # 2. 建立「汇总 Dashboard」 Sheet (放在第一页)
    # -------------------------------------------------------------
    ws_sum = wb.active
    ws_sum.title = "汇总 Dashboard"
    ws_sum.views.sheetView[0].showGridLines = True

    periods = list(summary_by_period.keys())
    sum_headers = ["持卡人"] + periods + ["合计 (Total)"]

    # 写入表头 (第一列已更改为「持卡人」)
    for c_idx, h_text in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=1, column=c_idx, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # 获取所有持卡人名字
    all_holders = sorted(list({h for p in summary_by_period.values() for h in p.keys()}))

    # 写入静态数值数据行
    for r_idx, holder in enumerate(all_holders, start=2):
        c1 = ws_sum.cell(row=r_idx, column=1, value=holder)
        c1.font = data_font
        c1.alignment = Alignment(horizontal="left", vertical="center")
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
            
        # 静态计算出的整行合计
        c_tot = ws_sum.cell(row=r_idx, column=len(periods) + 2, value=row_tot)
        c_tot.font = data_font
        c_tot.number_format = "$#,##0.00"
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_tot.border = thin_border

    # 写入底部合计 Total 行 (静态数值)
    tot_row = len(all_holders) + 2
    lbl = ws_sum.cell(row=tot_row, column=1, value="合计 (Total)")
    lbl.font = bold_font
    lbl.alignment = Alignment(horizontal="left", vertical="center")
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
    # 3. 增加 Check 校验行 (校验明细表 Charge 金额与首页 Total 单元格)
    # -------------------------------------------------------------
    check_row = tot_row + 2
    ws_sum.cell(row=check_row, column=1, value="check").font = bold_font
    
    # 动态构建 SUMIF 公式，相减目标单元格为 Grand Total 所在位置 (如 D6)
    c_check = ws_sum.cell(row=check_row, column=len(periods) + 2)
    c_check.value = f'=SUMIF(交易明细!J2:J{last_detail_row}, "Charge", 交易明细!I2:I{last_detail_row}) - {total_cell_letter}{tot_row}'
    c_check.font = bold_font
    # 会计格式：当完全相等相减为 0 时自动显示为 "-"
    c_check.number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    c_check.alignment = Alignment(horizontal="right", vertical="center")

    # 自动设置首页列宽
    ws_sum.column_dimensions['A'].width = 25
    for col_i in range(2, len(periods) + 3):
        col_l = get_column_letter(col_i)
        ws_sum.column_dimensions[col_l].width = 30

    wb.save(output_filename)
    print(f"导出成功！已保存文件为: {output_filename}")

# -------------------------------------------------------------
# 运行主程序示例
# -------------------------------------------------------------
if __name__ == "__main__":
    # 替换为您的 PDF 账单路径列表
    pdf_files = [
        "2273-Statement-20260715.pdf",
        "2273-Statement-20260731.pdf"
    ]
    
    # 1. 提取 PDF 账单数据
    summary_data, detail_data = extract_pdf_data(pdf_files)
    
    # 2. 生成 Excel 报告
    generate_excel_report(summary_data, detail_data, "bank_statement_final.xlsx")

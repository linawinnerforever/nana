import pandas as pd
import numpy as np
import datetime
import calendar
import streamlit as st
import io

def build_kingdee_voucher(draft_file_obj, date_str):
    """
    将上传的拆分底稿自动转换为金蝶上传凭证模板内存文件
    """
    # 1. 解析日期参数
    voucher_date = datetime.datetime.strptime(str(date_str), "%Y-%m-%d")
    year = voucher_date.year
    month = voucher_date.month
    
    # 获取当月最后一天
    _, last_day = calendar.monthrange(year, month)
    date_formatted = f"{year}-{month:02d}-{last_day:02d}"
    explanation = f"计提{year}年{month}月Claude消耗-主营业务成本_软件服务费"

    # 2. 读取底稿原始数据
    df_raw = pd.read_excel(draft_file_obj, header=None)
    
    project_segment_row = df_raw.iloc[0].values
    
    # 识别列头中有有效项目编码（如 001, 002 等）的列
    data_cols = []
    for col_idx in range(len(project_segment_row)):
        val = str(project_segment_row[col_idx]).strip()
        if val != 'nan' and val != '' and val.isdigit():
            data_cols.append((col_idx, val.zfill(3)))
            
    # 查找成本中心数据行（成本中心编码在第1列，从第3行开始）
    rows_data = []
    for r in range(3, len(df_raw)):
        dept_code = df_raw.iloc[r, 1]
        if pd.isna(dept_code):
            continue
        dept_code_str = str(int(dept_code)) if isinstance(dept_code, (int, float)) else str(dept_code).strip()
        
        for col_idx, proj_code in data_cols:
            amount = df_raw.iloc[r, col_idx]
            if pd.notna(amount):
                try:
                    amount_val = round(float(amount), 2)
                    if amount_val != 0:
                        rows_data.append({
                            'dept_code': dept_code_str,
                            'proj_code': proj_code,
                            'amount': amount_val
                        })
                except ValueError:
                    continue

    # 3. 构造金蝶凭证表头与结构
    header_row0 = [
        'FBillHead(GL_VOUCHER)', 'FAccountBookID', 'FAccountBookID#Name', 'FDate', 'FBUSDATE', 'FYEAR', 'FPERIOD',
        'FVOUCHERGROUPID', 'FVOUCHERGROUPID#Name', 'FVOUCHERGROUPNO', 'FATTACHMENTS', 'FISADJUSTVOUCHER',
        'FACCBOOKORGID', 'FACCBOOKORGID#Name', 'FSourceBillKey', 'FSourceBillKey#Name', 'FIMPORTVERSION',
        '*Split*1', 'FEntity', 'FEXPLANATION', 'FACCOUNTID', 'FACCOUNTID#Name', 'FDetailID#FF100003',
        'FDetailID#FF100003#Name', 'FDetailID#FF100002', 'FDetailID#FF100002#Name', 'FDetailID#FFLEX16',
        'FDetailID#FFLEX16#Name', 'FDetailID#FFLEX15', 'FDetailID#FFLEX15#Name', 'FDetailID#FFLEX14',
        'FDetailID#FFLEX14#Name', 'FDetailID#FFLEX13', 'FDetailID#FFLEX13#Name', 'FDetailID#FFLEX12',
        'FDetailID#FFLEX12#Name', 'FDetailID#FFLEX11', 'FDetailID#FFLEX11#Name', 'FDetailID#FFlex10',
        'FDetailID#FFlex10#Name', 'FDetailID#FFLEX9', 'FDetailID#FFLEX9#Name', 'FDetailID#FFlex8',
        'FDetailID#FFlex8#Name', 'FDetailID#FFlex7', 'FDetailID#FFlex7#Name', 'FDetailID#FFlex6',
        'FDetailID#FFlex6#Name', 'FDetailID#FFlex5', 'FDetailID#FFlex5#Name', 'FDetailID#FFlex4',
        'FDetailID#FFlex4#Name', 'FDetailID#FF100004', 'FDetailID#FF100004#Name', 'FDetailID#FF100005',
        'FDetailID#FF100005#Name', 'FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name',
        'FEXCHANGERATE', 'FUnitId', 'FUnitId#Name', 'FPrice', 'FQty', 'FAMOUNTFOR', 'FDEBIT', 'FCREDIT',
        'FSettleTypeID', 'FSettleTypeID#Name', 'FSETTLENO', 'FBUSNO', 'FEXPORTENTRYID'
    ]
    
    header_row1 = [
        '*单据头(序号)', '*(单据头)账簿#编码', '(单据头)账簿#名称', '*(单据头)日期', '(单据头)业务日期', '(单据头)会计年度',
        '(单据头)期间', '*(单据头)凭证字#编码', '(单据头)凭证字#名称', '*(单据头)凭证号', '(单据头)附件数', '(单据头)是否调整期凭证',
        '*(单据头)核算组织#编码', '(单据头)核算组织#名称', '(单据头)业务类型#编码', '(单据头)业务类型#名称', '(单据头)引入版本号',
        '间隔列', '*分录(序号)', '(分录)摘要', '*(分录)科目编码#编码', '(分录)科目编码#名称', '(分录)北京公司项目名#编码',
        '(分录)北京公司项目名#名称(Null)', '(分录)项目段#编码', '(分录)项目段#名称(Null)', '(分录)其他往来单位#编码',
        '(分录)其他往来单位#名称(Null)', '(分录)银行账号#编码', '(分录)银行账号#名称(Null)', '(分录)银行#编码',
        '(分录)银行#名称(Null)', '(分录)客户分组#编码', '(分录)客户分组#名称(Null)', '(分录)物料分组#编码',
        '(分录)物料分组#名称(Null)', '(分录)组织机构#编码', '(分录)组织机构#名称(Null)', '(分录)资产类别#编码',
        '(分录)资产类别#名称(Null)', '(分录)费用项目#编码', '(分录)费用项目#名称(Null)', '(分录)物料#编码',
        '(分录)物料#名称(Null)', '(分录)员工#编码', '(分录)员工#名称(Null)', '(分录)客户#编码',
        '(分录)客户#名称(Null)', '(分录)部门#编码', '(分录)部门#名称(Null)', '(分录)供应商#编码',
        '(分录)供应商#名称(Null)', '(分录)NL剧集#编码', '(分录)NL剧集#名称(Null)', '(分录)海南剧集#编码',
        '(分录)海南剧集#名称(Null)', '*(分录)币别#编码', '(分录)币别#名称', '*(分录)汇率类型#编码',
        '(分录)汇率类型#名称', '(分录)汇率', '(分录)单位#编码', '(分录)单位#名称', '(分录)单价', '(分录)数量',
        '(分录)原币金额', '(分录)借方金额', '(分录)贷方金额', '(分录)结算方式#编码', '(分录)结算方式#名称', '(分录)结算号',
        '(分录)业务编号', '(分录)现金流量#分录ID'
    ]

    result_rows = [header_row0, header_row1]
    total_debit_amount = 0.0
    entry_seq = 1

    # 4. 生成借方分录 (科目 6401.21)
    for item in rows_data:
        amt = item['amount']
        total_debit_amount += amt
        is_first = (entry_seq == 1)
        row = [
            '1' if is_first else np.nan,            # 0: *单据头(序号) -> 明确文本格式 '1'
            '002' if is_first else np.nan,          # 1: 账簿编码
            np.nan,                                 # 2: 账簿名称
            date_formatted if is_first else np.nan, # 3: 日期
            date_formatted if is_first else np.nan, # 4: 业务日期
            year if is_first else np.nan,           # 5: 会计年度
            month if is_first else np.nan,          # 6: 期间
            'PRE001' if is_first else np.nan,       # 7: 凭证字编码
            np.nan,                                 # 8: 凭证字名称
            1 if is_first else np.nan,              # 9: 凭证号
            np.nan, np.nan,                         # 10, 11
            '100' if is_first else np.nan,          # 12: 核算组织编码 -> 明确文本格式 '100'
            np.nan, np.nan, np.nan, np.nan, np.nan, # 13-17
            entry_seq,                              # 18: *分录(序号)
            explanation,                            # 19: 摘要
            '6401.21',                              # 20: *(分录)科目编码#编码
            np.nan, np.nan, np.nan,                 # 21-23
            item['proj_code'],                      # 24: FDetailID#FF100002 项目段#编码
            np.nan, np.nan, np.nan, np.nan, np.nan, # 25-29
            np.nan, np.nan, np.nan, np.nan, np.nan, # 30-34
            np.nan, np.nan, np.nan, np.nan, np.nan, # 35-39
            np.nan, np.nan, np.nan, np.nan, np.nan, # 40-44
            np.nan, np.nan, np.nan,                 # 45-47
            item['dept_code'],                      # 48: FDetailID#FFlex5 (分录)部门#编码
            np.nan,                                 # 49
            'VEN02027',                             # 50: FDetailID#FFlex4 (分录)供应商#编码
            np.nan, np.nan, np.nan, np.nan, np.nan, # 51-55
            'PRE007',                               # 56: FCURRENCYID *(分录)币别#编码
            '美元',                                 # 57: FCURRENCYID#Name (分录)币别#名称
            'HLTX01_SYS',                           # 58: FEXCHANGERATETYPE *(分录)汇率类型#编码
            '固定汇率',                             # 59: FEXCHANGERATETYPE#Name (分录)汇率类型#名称
            1,                                      # 60: FEXCHANGERATE (分录)汇率
            np.nan, np.nan, np.nan, np.nan,         # 61-64
            np.nan,                                 # 65: FAMOUNTFOR 原币金额 (留空)
            amt,                                    # 66: FDEBIT 借方金额
            np.nan,                                 # 67: FCREDIT 贷方金额 (为空)
            np.nan, np.nan, np.nan, np.nan, np.nan  # 68-72
        ]
        result_rows.append(row)
        entry_seq += 1

    # 5. 生成贷方分录 (科目 2202.01)
    total_debit_amount = round(total_debit_amount, 2)
    credit_row = [
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        entry_seq,                                  # 18: 分录序号
        explanation,                                # 19: 摘要
        '2202.01',                                  # 20: 贷方科目编码
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        'VEN02027',                                 # 50: FDetailID#FFlex4 供应商编码
        np.nan, np.nan, np.nan, np.nan, np.nan,
        'PRE007',                                   # 56: 币别编码
        '美元',                                     # 57: 币别名称
        'HLTX01_SYS',                               # 58: 汇率类型编码
        '固定汇率',                                 # 59: 汇率类型名称
        1,                                          # 60: 汇率
        np.nan, np.nan, np.nan, np.nan,
        np.nan,                                     # 65: FAMOUNTFOR 原币金额 (留空)
        np.nan,                                     # 66: 借方金额 (为空)
        total_debit_amount,                         # 67: FCREDIT 贷方金额
        np.nan, np.nan, np.nan, np.nan, np.nan
    ]
    result_rows.append(credit_row)

    # 6. 将生成的 Excel 写入内存数据流中，并强制设定单元格文本格式
    out_df = pd.DataFrame(result_rows)
    output_stream = io.BytesIO()
    
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        out_df.to_excel(writer, index=False, header=False, sheet_name='凭证#单据头(FBillHead)')
        worksheet = writer.sheets['凭证#单据头(FBillHead)']
        
        # 强制设置整列或单元格为文本格式 (Format '@')
        # Col 1: FBillHead(GL_VOUCHER) (列索引 A)
        # Col 13: FACCBOOKORGID (列索引 M)
        for row in range(3, len(result_rows) + 1):
            cell_billhead = worksheet.cell(row=row, column=1)
            cell_orgid = worksheet.cell(row=row, column=13)
            
            cell_billhead.number_format = '@'
            cell_orgid.number_format = '@'
            
            if cell_billhead.value is not None:
                cell_billhead.value = str(cell_billhead.value)
            if cell_orgid.value is not None:
                cell_orgid.value = str(cell_orgid.value)

    output_stream.seek(0)
    return output_stream, entry_seq, total_debit_amount

# --- Streamlit 网页前端界面 ---
st.title("📊 Claude金蝶入账凭证生成工具")

uploaded_file = st.file_uploader("请选择或拖入当月【Claude拆分底稿】Excel 文件", type=["xlsx", "xls"])
voucher_date_input = st.date_input("凭证日期", value=datetime.date(2026, 7, 31))

if uploaded_file is not None:
    if st.button("🚀 开始自动转换"):
        try:
            excel_data, total_entries, total_amt = build_kingdee_voucher(uploaded_file, voucher_date_input)
            st.success(f"转换成功！生成凭证分录 {total_entries} 条，凭证总金额 ${total_amt:,.2f}")
            
            file_name = f"金蝶上传凭证_{voucher_date_input.strftime('%Y%m')}.xlsx"
            st.download_button(
                label="📥 点击下载金蝶凭证上传文件",
                data=excel_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"转换过程出现错误：{str(e)}")

import pandas as pd
import numpy as np
import datetime
import calendar
import streamlit as st
import io

def build_single_vendor_voucher(sub_df, data_cols, date_formatted, year, month, vendor_name_text):
    """
    针对单家供应商（OpenAI 或 Google）数据生成标准金蝶凭证
    """
    month_abbr = calendar.month_abbr[month]
    fixed_summary = f"计提{year}年{month}月服务器成本-{month_abbr}.{year} server cost accrual-{vendor_name_text}"
    
    parsed_rows = []
    
    for idx, row in sub_df.iterrows():
        dept_code = row[5] # Col 5: 成本中心编码
        if pd.isna(dept_code) or str(dept_code).strip() in ['nan', '']:
            continue
            
        dept_code_str = str(int(dept_code)) if isinstance(dept_code, (int, float)) else str(dept_code).strip()
        vendor_code = str(row[10]).strip() if pd.notna(row[10]) else ""
        account_code = str(row[11]).strip() if pd.notna(row[11]) else ""
        
        for col_idx, proj_code in data_cols:
            amt = row[col_idx]
            if pd.notna(amt):
                try:
                    amt_val = round(float(amt), 2)
                    if amt_val != 0:
                        parsed_rows.append({
                            'dept_code': dept_code_str,
                            'proj_code': proj_code,
                            'vendor_code': vendor_code,
                            'account_code': account_code,
                            'summary': fixed_summary,
                            'amount': amt_val
                        })
                except ValueError:
                    continue

    if not parsed_rows:
        return None, 0, 0.0

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
    last_vendor = ""

    for item in parsed_rows:
        amt = item['amount']
        total_debit_amount += amt
        is_first = (entry_seq == 1)
        
        hainan_code = '025' if item['account_code'] == '6401.03.04' else np.nan
        last_vendor = item['vendor_code']

        row = [
            '1' if is_first else np.nan,            # 0: *单据头(序号)
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
            '100' if is_first else np.nan,          # 12: 核算组织编码
            np.nan, np.nan, np.nan, np.nan, np.nan, # 13-17
            entry_seq,                              # 18: *分录(序号)
            item['summary'],                        # 19: 摘要
            item['account_code'],                   # 20: *(分录)科目编码#编码
            np.nan, np.nan, np.nan,                 # 21-23
            item['proj_code'],                      # 24: 项目段#编码
            np.nan, np.nan, np.nan, np.nan, np.nan, # 25-29
            np.nan, np.nan, np.nan, np.nan, np.nan, # 30-34
            np.nan, np.nan, np.nan, np.nan, np.nan, # 35-39
            np.nan, np.nan, np.nan, np.nan, np.nan, # 40-44
            np.nan, np.nan, np.nan,                 # 45-47
            item['dept_code'],                      # 48: 部门编码
            np.nan,                                 # 49
            item['vendor_code'],                    # 50: 供应商编码
            np.nan, np.nan, np.nan,                 # 51-53
            hainan_code,                            # 54: 海南剧集#编码
            np.nan,                                 # 55
            'PRE007', '美元', 'HLTX01_SYS', '固定汇率', 1,
            np.nan, np.nan, np.nan, np.nan,
            np.nan,                                 # 65: 原币金额
            amt,                                    # 66: 借方金额
            np.nan,                                 # 67: 贷方金额
            np.nan, np.nan, np.nan, np.nan, np.nan  # 68-72
        ]
        result_rows.append(row)
        entry_seq += 1

    total_debit_amount = round(total_debit_amount, 2)
    credit_row = [
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        entry_seq, fixed_summary, '2202.01',
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        last_vendor, np.nan, np.nan, np.nan, np.nan, np.nan,
        'PRE007', '美元', 'HLTX01_SYS', '固定汇率', 1,
        np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, total_debit_amount,
        np.nan, np.nan, np.nan, np.nan, np.nan
    ]
    result_rows.append(credit_row)

    out_df = pd.DataFrame(result_rows)
    output_stream = io.BytesIO()
    sheet_target = '凭证#单据头(FBillHead)'
    
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        out_df.to_excel(writer, index=False, header=False, sheet_name=sheet_target)
        worksheet = writer.sheets[sheet_target]
        
        for row in range(3, len(result_rows) + 1):
            cell_billhead = worksheet.cell(row=row, column=1)
            cell_orgid = worksheet.cell(row=row, column=13)
            cell_billhead.number_format = '@'
            cell_orgid.number_format = '@'
            if cell_billhead.value is not None:
                cell_billhead.value = str(cell_billhead.value)
            if cell_orgid.value is not None:
                cell_orgid.value = str(cell_orgid.value)

    output_stream.getvalue() # Ensure stream data
    return output_stream.getvalue(), entry_seq, total_debit_amount


def process_server_cost_sheet(draft_file_obj, date_str):
    voucher_date = datetime.datetime.strptime(str(date_str), "%Y-%m-%d")
    year = voucher_date.year
    month = voucher_date.month
    
    _, last_day = calendar.monthrange(year, month)
    date_formatted = f"{year}-{month:02d}-{last_day:02d}"

    df_raw = pd.read_excel(draft_file_obj, sheet_name=0, header=None)
    
    proj_row = df_raw.iloc[0].values
    data_cols = []
    for c in range(14, 21):
        v = str(proj_row[c]).strip()
        if v != 'nan' and v != '' and v.isdigit():
            data_cols.append((c, v.zfill(3)))

    openai_sub_df = df_raw[df_raw.iloc[:, 2] == 'OpenAI']
    google_sub_df = df_raw[df_raw.iloc[:, 2] == 'Google cloud']

    openai_data, openai_entries, openai_amt = build_single_vendor_voucher(
        openai_sub_df, data_cols, date_formatted, year, month, vendor_name_text="OpenAI, LLC"
    )
    
    google_data, google_entries, google_amt = build_single_vendor_voucher(
        google_sub_df, data_cols, date_formatted, year, month, vendor_name_text="Google Cloud"
    )

    return {
        'OpenAI': (openai_data, openai_entries, openai_amt),
        'Google': (google_data, google_entries, google_amt)
    }

# --- Streamlit 网页前端界面 ---
st.title("📊 OpenAI & Google金蝶入账凭证生成工具")

uploaded_file = st.file_uploader("请选择或拖入当月【测试服务器成本分摊汇总表】Excel 文件", type=["xlsx", "xls"])
voucher_date_input = st.date_input("凭证日期", value=datetime.date(2026, 7, 31))

# 初始化 Session State，用于持久化保存生成结果
if 'generated_results' not in st.session_state:
    st.session_state.generated_results = None

# 当用户点击“生成凭证”时，触发数据解析并存入 session_state
if uploaded_file is not None:
    if st.button("🚀 开始自动转换 (生成 OpenAI 与 Google 凭证)"):
        with st.spinner("正在解析底稿并生成凭证..."):
            try:
                st.session_state.generated_results = process_server_cost_sheet(uploaded_file, voucher_date_input)
                st.session_state.date_tag = voucher_date_input.strftime('%Y%m')
            except Exception as e:
                st.error(f"转换过程出现错误：{str(e)}")

# 渲染生成结果区域（只要 session_state 有数据，点击下载就不丢失）
if st.session_state.generated_results is not None:
    st.subheader("生成结果：")
    date_tag = st.session_state.get('date_tag', '202607')
    results = st.session_state.generated_results
    
    # OpenAI 下载区
    openai_data, openai_entries, openai_amt = results['OpenAI']
    if openai_data:
        st.success(f"✅ OpenAI 凭证生成成功：共 {openai_entries} 条分录，总金额 ${openai_amt:,.2f}")
        st.download_button(
            label="📥 点击下载【OpenAI 金蝶凭证】",
            data=openai_data,
            file_name=f"OpenAI_金蝶上传凭证_{date_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_download_openai"
        )
    
    # Google 下载区
    google_data, google_entries, google_amt = results['Google']
    if google_data:
        st.success(f"✅ Google 凭证生成成功：共 {google_entries} 条分录，总金额 ${google_amt:,.2f}")
        st.download_button(
            label="📥 点击下载【Google 金蝶凭证】",
            data=google_data,
            file_name=f"Google_金蝶上传凭证_{date_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_download_google"
        )

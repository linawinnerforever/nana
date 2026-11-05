import pandas as pd
import numpy as np
import datetime
import calendar
import streamlit as st
import io

def build_single_vendor_voucher(sub_df, data_cols, date_formatted, year, month, summary_text, credit_account):
    """
    针对单家供应商/产品数据生成标准金蝶凭证
    """
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
                            'summary': summary_text,
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

    # 生成借方分录
    for item in parsed_rows:
        amt = item['amount']
        total_debit_amount += amt
        is_first = (entry_seq == 1)
        
        # 针对 6401.03.04 科目自动匹配海南剧集编码 025
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
            'PRE007', '美元', 'HLTX01_SYS', '固定汇率', 1, # 56-60
            np.nan, np.nan, np.nan, np.nan,         # 61-64
            np.nan,                                 # 65: 原币金额 (留空)
            amt,                                    # 66: 借方金额
            np.nan,                                 # 67: 贷方金额
            np.nan, np.nan, np.nan, np.nan, np.nan  # 68-72
        ]
        result_rows.append(row)
        entry_seq += 1

    # 生成贷方分录 (Cursor 贷方为 1123.04，其余为 2202.01)
    total_debit_amount = round(total_debit_amount, 2)
    credit_row = [
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        entry_seq,                                  # 18: 分录序号
        summary_text,                               # 19: 摘要
        credit_account,                             # 20: 贷方科目编码
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        last_vendor,                                # 50: 供应商编码
        np.nan, np.nan, np.nan, np.nan, np.nan,
        'PRE007', '美元', 'HLTX01_SYS', '固定汇率', 1,
        np.nan, np.nan, np.nan, np.nan,
        np.nan,                                     # 65: 原币金额 (留空)
        np.nan,                                     # 66: 借方金额
        total_debit_amount,                         # 67: 贷方金额
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

    output_stream.getvalue()
    return output_stream.getvalue(), entry_seq, total_debit_amount


def process_server_cost_sheet(draft_file_obj, date_str):
    """
    读取分摊底稿，生成 Google Cloud、OpenAI、Cursor、Claude 四张凭证
    """
    voucher_date = datetime.datetime.strptime(str(date_str), "%Y-%m-%d")
    year = voucher_date.year
    month = voucher_date.month
    month_abbr = calendar.month_abbr[month] # 如 Oct, Nov
    
    _, last_day = calendar.monthrange(year, month)
    date_formatted = f"{year}-{month:02d}-{last_day:02d}"

    # 始终读取第一个工作表
    df_raw = pd.read_excel(draft_file_obj, sheet_name=0, header=None)
    
    # 1. 提取第 0 行的项目段编码列
    proj_row = df_raw.iloc[0].values
    data_cols = []
    for c in range(14, len(proj_row)):
        v = str(proj_row[c]).strip()
        if v != 'nan' and v != '' and v.isdigit():
            data_cols.append((c, v.zfill(3)))

    # 2. 动态定位真正的表头行（包含“成本中心编码”和“供应商”），严格排除顶部汇总区域
    header_row_idx = 5 # 默认兜底为第5行
    for r in range(min(15, len(df_raw))):
        row_str = " ".join([str(x) for x in df_raw.iloc[r].values])
        if "成本中心编码" in row_str and "供应商" in row_str:
            header_row_idx = r
            break

    # 核心：明细数据必须从表头行的下一行开始提取
    df_data_only = df_raw.iloc[header_row_idx + 1:].copy()

    # 3. 4 款产品的专属配置
    products_config = {
        'Google': {
            'match_pattern': r'google',
            'summary': f"计提{year}年{month:02d}月服务器成本-{month_abbr}.{year} server cost accrual-Google Cloud",
            'credit_account': '2202.01',
            'title': 'Google Cloud'
        },
        'OpenAI': {
            'match_pattern': r'openai',
            'summary': f"计提{year}年{month:02d}月服务器成本-{month_abbr}.{year} server cost accrual-OpenAI, LLC",
            'credit_account': '2202.01',
            'title': 'OpenAI'
        },
        'Cursor': {
            'match_pattern': r'cursor',
            'summary': f"计提{year}年{month:02d}月Cursor消耗",
            'credit_account': '1123.04',
            'title': 'Cursor'
        },
        'Claude': {
            'match_pattern': r'claude',
            'summary': f"计提{year}年{month:02d}月Claude消耗",
            'credit_account': '2202.01',
            'title': 'Claude'
        }
    }

    vendor_col = df_data_only.iloc[:, 2].astype(str)

    results = {}
    for key, cfg in products_config.items():
        sub_df = df_data_only[vendor_col.str.contains(cfg['match_pattern'], case=False, na=False)]
        
        file_bytes, entries_cnt, total_amt = build_single_vendor_voucher(
            sub_df, 
            data_cols, 
            date_formatted, 
            year, 
            month, 
            summary_text=cfg['summary'], 
            credit_account=cfg['credit_account']
        )
        results[key] = {
            'data': file_bytes,
            'entries': entries_cnt,
            'total_amt': total_amt,
            'title': cfg['title'],
            'credit_account': cfg['credit_account']
        }

    return results

# --- Streamlit 网页前端界面 ---
st.title("📊 AI工具金蝶入账凭证生成工具")

uploaded_file = st.file_uploader("请选择或拖入当月【成本分摊汇总表】Excel 文件", type=["xlsx", "xls"])
voucher_date_input = st.date_input("凭证日期", value=datetime.date(2026, 10, 31))

# 持久化存储结果，解决下载按钮刷新问题
if 'multi_product_results' not in st.session_state:
    st.session_state.multi_product_results = None

if uploaded_file is not None:
    if st.button("🚀 开始自动转换 (同时生成 4 款产品凭证)"):
        with st.spinner("正在解析底稿并生成各产品凭证..."):
            try:
                st.session_state.multi_product_results = process_server_cost_sheet(uploaded_file, voucher_date_input)
                st.session_state.date_tag = voucher_date_input.strftime('%Y%m')
            except Exception as e:
                st.error(f"转换过程出现错误：{str(e)}")

# 渲染生成结果区域
if st.session_state.multi_product_results is not None:
    st.subheader("生成结果：")
    date_tag = st.session_state.get('date_tag', voucher_date_input.strftime('%Y%m'))
    results = st.session_state.multi_product_results
    
    for key, res in results.items():
        if res['data'] is not None:
            st.success(f"✅ **{res['title']}** 凭证生成成功：共 {res['entries']} 条分录，总借贷金额 ${res['total_amt']:,.2f}（贷方科目: `{res['credit_account']}`）")
            st.download_button(
                label=f"📥 点击下载【{res['title']} 金蝶凭证】",
                data=res['data'],
                file_name=f"{key}_金蝶上传凭证_{date_tag}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"btn_download_{key}"
            )
        else:
            st.warning(f"⚠️ 未在底稿中检测到 **{res['title']}** 的有效拆分数据。")

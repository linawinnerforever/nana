import pandas as pd
import numpy as np
import datetime
import calendar
import streamlit as st
import io

def build_server_summary_voucher(draft_file_obj, date_str):
    """
    读取《服务器成本分摊汇总表》中的“汇总”页签，自动打包生成标准金蝶上传凭证 Excel
    """
    # 1. 解析日期参数
    voucher_date = datetime.datetime.strptime(str(date_str), "%Y-%m-%d")
    year = voucher_date.year
    month = voucher_date.month
    
    _, last_day = calendar.monthrange(year, month)
    date_formatted = f"{year}-{month:02d}-{last_day:02d}"

    # 2. 读取“汇总”Sheet 原始数据
    xls = pd.ExcelFile(draft_file_obj)
    sheet_name_src = '汇总' if '汇总' in xls.sheet_names else xls.sheet_names[0]
    df_summary = pd.read_excel(xls, sheet_name=sheet_name_src, header=None)

    # 第 10 行为表头，第 11 行起为明细数据
    df_data = df_summary.iloc[11:].copy()
    df_data.columns = df_summary.iloc[10].values

    # 过滤入账美元金额为 0 或空的数据
    df_data['求和项:入账美元金额'] = pd.to_numeric(df_data['求和项:入账美元金额'], errors='coerce').fillna(0)
    df_valid = df_data[df_data['求和项:入账美元金额'] != 0].copy()

    # 保持原有出现顺序按 (摘要-final, 供应商编码) 进行分组
    grouped_keys = []
    for idx, row in df_valid.iterrows():
        summary_val = str(row['摘要-final']).strip() if pd.notna(row['摘要-final']) else ""
        vendor_val = str(row['供应商编码']).strip() if pd.notna(row['供应商编码']) else ""
        key = (summary_val, vendor_val)
        if key not in grouped_keys:
            grouped_keys.append(key)

    # 3. 构造金蝶凭证表头结构
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
    total_debit_grand_sum = 0.0
    entry_seq = 1

    # 4. 生成分录行
    for summary_val, vendor_code in grouped_keys:
        sub_group = df_valid[
            (df_valid['摘要-final'].astype(str).str.strip() == summary_val) & 
            (df_valid['供应商编码'].astype(str).str.strip() == vendor_code)
        ]
        
        group_debit_total = 0.0
        
        # 4.1 逐行生成借方分录
        for idx, row in sub_group.iterrows():
            amt = round(float(row['求和项:入账美元金额']), 2)
            group_debit_total += amt
            total_debit_grand_sum += amt
            
            account_code = str(row['科目编码']).strip() if pd.notna(row['科目编码']) else ""
            
            proj_code = str(row['项目编码']).strip() if pd.notna(row['项目编码']) else ""
            if proj_code.isdigit():
                proj_code = proj_code.zfill(3)
            elif proj_code in ['nan', 'None']:
                proj_code = ""

            dept_code = str(row['部门编码']).strip() if ('部门编码' in row and pd.notna(row['部门编码'])) else np.nan
            if str(dept_code) in ['nan', 'None', '']:
                dept_code = np.nan

            hainan_code = '025' if account_code == '6401.03.04' else np.nan
            is_first = (entry_seq == 1)

            debit_row = [
                '1' if is_first else np.nan,            # 0: *单据头(序号)
                '002' if is_first else np.nan,          # 1: 账簿编码
                'Crazy Maple Studio Inc' if is_first else np.nan, # 2: 账簿名称
                date_formatted if is_first else np.nan, # 3: 日期
                date_formatted if is_first else np.nan, # 4: 业务日期
                year if is_first else np.nan,           # 5: 会计年度
                month if is_first else np.nan,          # 6: 期间
                'PRE001' if is_first else np.nan,       # 7: 凭证字编码
                '记' if is_first else np.nan,           # 8: 凭证字名称
                1 if is_first else np.nan,              # 9: 凭证号
                np.nan, np.nan,                         # 10, 11
                '100' if is_first else np.nan,          # 12: 核算组织编码
                np.nan, np.nan, np.nan, np.nan, np.nan, # 13-17
                entry_seq,                              # 18: *分录(序号)
                summary_val,                            # 19: 摘要
                account_code,                           # 20: 科目编码
                np.nan, np.nan, np.nan,                 # 21-23
                proj_code,                              # 24: 项目段编码
                np.nan, np.nan, np.nan, np.nan, np.nan, # 25-29
                np.nan, np.nan, np.nan, np.nan, np.nan, # 30-34
                np.nan, np.nan, np.nan, np.nan, np.nan, # 35-39
                np.nan, np.nan, np.nan, np.nan, np.nan, # 40-44
                np.nan, np.nan, np.nan,                 # 45-47
                dept_code,                              # 48: 部门编码
                np.nan,                                 # 49
                vendor_code,                            # 50: 供应商编码
                np.nan, np.nan, np.nan,                 # 51-53
                hainan_code,                            # 54: 海南剧集编码
                np.nan,                                 # 55
                'PRE007', np.nan, 'HLTX01_SYS', np.nan, 1, # 56-60
                np.nan, np.nan, np.nan, np.nan,         # 61-64
                np.nan,                                 # 65: 原币金额
                amt,                                    # 66: 借方金额
                np.nan,                                 # 67: 贷方金额
                np.nan, np.nan, np.nan, np.nan, np.nan  # 68-72
            ]
            result_rows.append(debit_row)
            entry_seq += 1

        # 4.2 生成该组对应的贷方分录 (应付账款 2202.01)
        group_debit_total = round(group_debit_total, 2)
        credit_row = [
            np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
            np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
            entry_seq,                                  # 18: 分录序号
            summary_val,                                # 19: 摘要
            '2202.01',                                  # 20: 贷方科目编码
            np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
            np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
            np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
            vendor_code,                                # 50: 供应商编码
            np.nan, np.nan, np.nan, np.nan, np.nan,
            'PRE007', np.nan, 'HLTX01_SYS', np.nan, 1,
            np.nan, np.nan, np.nan, np.nan,
            np.nan,                                     # 65: 原币金额
            np.nan,                                     # 66: 借方金额
            group_debit_total,                          # 67: 贷方金额
            np.nan, np.nan, np.nan, np.nan, np.nan
        ]
        result_rows.append(credit_row)
        entry_seq += 1

    # 5. 生成 Excel 并设置工作表名称和格式
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
    return output_stream.getvalue(), entry_seq - 1, round(total_debit_grand_sum, 2)

# --- Streamlit 网页前端界面 ---
st.title("📊 服务器成本分摊汇总凭证生成工具")

uploaded_file = st.file_uploader("请选择或拖入当月【测试服务器成本分摊汇总表】Excel 文件", type=["xlsx", "xls"])
voucher_date_input = st.date_input("凭证日期", value=datetime.date(2026, 7, 31))

# 持久化存储结果，解决下载按钮跳转问题
if 'server_summary_results' not in st.session_state:
    st.session_state.server_summary_results = None

if uploaded_file is not None:
    if st.button("🚀 开始自动生成金蝶凭证"):
        with st.spinner("正在读取汇总表并生成标准凭证..."):
            try:
                data_bytes, entries_cnt, total_amt = build_server_summary_voucher(uploaded_file, voucher_date_input)
                st.session_state.server_summary_results = {
                    'data': data_bytes,
                    'entries': entries_cnt,
                    'total_amt': total_amt,
                    'date_tag': voucher_date_input.strftime('%Y%m')
                }
            except Exception as e:
                st.error(f"解析发生错误：{str(e)}")

# 渲染持久化的结果与下载按钮
if st.session_state.server_summary_results is not None:
    res = st.session_state.server_summary_results
    st.subheader("生成结果：")
    st.success(f"✅ 凭证生成成功：共 {res['entries']} 条分录，总借贷金额 ${res['total_amt']:,.2f}")
    
    st.download_button(
        label="📥 点击下载【服务器成本金蝶上传凭证】",
        data=res['data'],
        file_name=f"服务器成本金蝶上传凭证_{res['date_tag']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_server_summary"
    )

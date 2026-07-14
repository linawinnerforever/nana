# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="金蝶云星空-集团费重分类全自动工具", layout="wide")

st.title("📊 金蝶云星空 - 费用重分类集团全自动生成凭证工具")

# ----------------------------------------------------
# 集团 18 大主体真实财务静态配置库
# ----------------------------------------------------
COMPANY_CONFIG = {
    "Crazy Maple Studio Inc": {"book_id": "002", "org_id": "100", "currency_id": "PRE007", "currency_name": "美元"},
    "CRAZY MAPLE  SERVICE  COMPANY": {"book_id": "003", "org_id": "101", "currency_id": "PRE007", "currency_name": "美元"},
    "Crazy Maple  Canada  Inc": {"book_id": "004", "org_id": "102", "currency_id": "PRE008", "currency_name": "加币"},
    "Maple House Inc": {"book_id": "005", "org_id": "103", "currency_id": "PRE007", "currency_name": "美元"},
    "CRAZY MAPLE  INTERACTIVE  HOLDING LTD": {"book_id": "006", "org_id": "104", "currency_id": "PRE007", "currency_name": "美元"},
    "SPICY MAPLE  LIMITED": {"book_id": "007", "org_id": "105", "currency_id": "PRE007", "currency_name": "美元"},
    "CRAZY MAPLE  STUDIO  HK LIMITED": {"book_id": "008", "org_id": "106", "currency_id": "PRE002", "currency_name": "香港元"},
    "New Leaf  Publishing  Inc": {"book_id": "009", "org_id": "107", "currency_id": "PRE007", "currency_name": "美元"},
    "北京枫悦互动科技有限公司": {"book_id": "010", "org_id": "108", "currency_id": "PRE001", "currency_name": "人民币"},
    "深圳枫叶互动科技有限公司": {"book_id": "011", "org_id": "109", "currency_id": "PRE001", "currency_name": "人民币"},
    "杭州枫叶互动科技有限公司": {"book_id": "012", "org_id": "110", "currency_id": "PRE001", "currency_name": "人民币"},
    "B25 LIMITED": {"book_id": "013", "org_id": "111", "currency_id": "PRE007", "currency_name": "美元"},
    "海南枫悦互动科技有限公司": {"book_id": "014", "org_id": "112", "currency_id": "PRE001", "currency_name": "人民币"},
    "北京枫悦互动科技有限公司工会委员会": {"book_id": "015", "org_id": "113", "currency_id": "PRE001", "currency_name": "人民币"},
    "深圳市星尘游戏科技有限公司": {"book_id": "016", "org_id": "114", "currency_id": "PRE001", "currency_name": "人民币"},
    "ReelShort Japan Co., Ltd.": {"book_id": "017", "org_id": "115", "currency_id": "PRE004", "currency_name": "日本日圆"},
    "SWEET MAPLE LIMITED": {"book_id": "018", "org_id": "116", "currency_id": "PRE007", "currency_name": "美元"},
    "深圳枫悦互动科技有限公司": {"book_id": "117", "org_id": "019", "currency_id": "PRE001", "currency_name": "人民币"}
}

CURRENCY_OPTIONS = {
    "人民币 (PRE001)": {"id": "PRE001", "name": "人民币"},
    "美元 (PRE007)": {"id": "PRE007", "name": "美元"},
    "香港元 (PRE002)": {"id": "PRE002", "name": "香港元"},
    "加币 (PRE008)": {"id": "PRE008", "name": "加币"},
    "日本日圆 (PRE004)": {"id": "PRE004", "name": "日本日圆"}
}

# ----------------------------------------------------
# 侧边栏控制台
# ----------------------------------------------------
st.sidebar.header("🛠️ 集团财务控制台")

company_options = list(COMPANY_CONFIG.keys())
selected_company = st.sidebar.selectbox("请选择本次做账的公司主体", options=company_options)
comp_info = COMPANY_CONFIG[selected_company]
st.sidebar.success(f"📍 锁定主体：账簿({comp_info['book_id']}) | 组织({comp_info['org_id']})")

default_curr_label = f"{comp_info['currency_name']} ({comp_info['currency_id']})"
if default_curr_label not in CURRENCY_OPTIONS:
    CURRENCY_OPTIONS[default_curr_label] = {"id": comp_info['currency_id'], "name": comp_info['currency_name']}

curr_labels = list(CURRENCY_OPTIONS.keys())
selected_curr_label = st.sidebar.selectbox("请选择记账本位币 (FCURRENCYID)", options=curr_labels, index=curr_labels.index(default_curr_label))
chosen_currency = CURRENCY_OPTIONS[selected_curr_label]

current_year = st.sidebar.number_input("会计年度 (FYEAR)", min_value=2020, max_value=2035, value=2026)
current_period = st.sidebar.slider("会计期间 (FPERIOD)", min_value=1, max_value=12, value=6)
last_day = calendar.monthrange(current_year, current_period)[1]
voucher_date = f"{current_year}-{str(current_period).zfill(2)}-{str(last_day).zfill(2)}"
st.sidebar.info(f"📅 凭证自动生成日期：{voucher_date}")

# ----------------------------------------------------
# 主界面上传区
# ----------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 上传费用明细表 (TB)")
    source_file = st.file_uploader("支持 .xlsx 格式", type=["xlsx"], key="src")
with col2:
    st.subheader("2. 上传部门项目分摊比例表")
    ratio_file = st.file_uploader("支持 .xlsx 格式", type=["xlsx"], key="ratio")

# 73列中英文标准矩阵牢牢锁死
tech_headers = [
    'FBillHead(GL_VOUCHER)', 'FAccountBookID', 'FAccountBookID#Name', 'FDate', 'FBUSDATE', 'FYEAR', 'FPERIOD', 
    'FVOUCHERGROUPID', 'FVOUCHERGROUPID#Name', 'FVOUCHERGROUPNO', 'FATTACHMENTS', 'FISADJUSTVOUCHER', 
    'FACCBOOKORGID', 'FACCBOOKORGID#Name', 'FSourceBillKey', 'FSourceBillKey#Name', 'FIMPORTVERSION', 
    '*Split*1', 'FEntity', 'FEXPLANATION', 'FACCOUNTID', 'FACCOUNTID#Name', 
    'FDetailID#FF100003', 'FDetailID#FF100003#Name', 'FDetailID#FF100002', 'FDetailID#FF100002#Name', 
    'FDetailID#FFLEX16', 'FDetailID#FFLEX16#Name', 'FDetailID#FFLEX15', 'FDetailID#FFLEX15#Name', 
    'FDetailID#FFLEX14', 'FDetailID#FFLEX14#Name', 'FDetailID#FFLEX13', 'FDetailID#FFLEX13#Name', 
    'FDetailID#FFLEX12', 'FDetailID#FFLEX12#Name', 'FDetailID#FFLEX11', 'FDetailID#FFLEX11#Name', 
    'FDetailID#FFlex10', 'FDetailID#FFlex10#Name', 'FDetailID#FFLEX9', 'FDetailID#FFLEX9#Name', 
    'FDetailID#FFlex8', 'FDetailID#FFlex8#Name', 'FDetailID#FFlex7', 'FDetailID#FFlex7#Name', 
    'FDetailID#FFlex6', 'FDetailID#FFlex6#Name', 'FDetailID#FFlex5', 'FDetailID#FFlex5#Name', 
    'FDetailID#FFlex4', 'FDetailID#FFlex4#Name', 'FDetailID#FF100004', 'FDetailID#FF100004#Name', 
    'FDetailID#FF100005', 'FDetailID#FF100005#Name', 
    'FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE', 
    'FUnitId', 'FUnitId#Name', 'FPrice', 'FQty', 'FAMOUNTFOR', 'FDEBIT', 'FCREDIT', 
    'FSettleTypeID', 'FSettleTypeID#Name', 'FSETTLENO', 'FBUSNO', 'FEXPORTENTRYID'
]

cn_headers = [
    '*单据头(序号)', '*(单据头)账簿#编码', '(单据头)账簿#名称', '*(单据头)日期', '(单据头)业务日期', 
    '(单据头)会计年度', '(单据头)期间', '*(单据头)凭证字#编码', '(单据头)凭证字#名称', 
    '*(单据头)凭证号', '(单据头)附件数', '(单据头)是否调整期凭证', '(单据头)核算组织#编码', 
    '(单据头)核算组织#名称', '(单据头)业务类型#编码', '(单据头)业务类型#名称', '(单据头)引入版本号', 
    '间隔列', '*分录(序号)', '(分录)摘要', '*(分录)科目编码#编码', '(分录)科目编码#名称', 
    '(分录)北京公司项目名#编码', '(分录)北京公司项目名#名称(Null)', '(分录)项目段#编码', '(分录)项目段#名称(Null)', 
    '(分录)其他往来单位#编码', '(分录)其他往来单位#名称(Null)', '(分录)银行账号#编码', '(分录)银行账号#名称(Null)', 
    '(分录)银行#编码', '(分录)银行#名称(Null)', '(分录)客户分组#编码', '(分录)客户分组#名称(Null)', 
    '(分录)物料分组#编码', '(分录)物料分组#名称(Null)', '(分录)组织机构#编码', '(分录)组织机构#名称(Null)', 
    '(分录)资产类别#编码', '(分录)资产类别#名称(Null)', '(分录)费用项目#编码', '(分录)费用项目#名称(Null)', 
    '(分录)物料#编码', '(分录)物料#名称(Null)', '(分录)国家/地区#编码', '(分录)国家/地区#名称(Null)', 
    '(分录)部门_海外#编码', '(分录)部门_海外#名称(Null)', '(分录)成本中心#编码', '(分录)成本中心#名称(Null)', 
    '(分录)职员#编码', '(分录)职员#名称(Null)', '(分录)客户#编码', '(分录)客户#名称(Null)', 
    '(分录)供应商#编码', '(分录)供应商#名称(Null)', 
    '(分录)币别#编码', '(分录)币别#名称', '(分录)汇率类型#编码', '(分录)汇率类型#名称', '(分录)汇率', 
    '(分录)计量单位#编码', '(分录)计量单位#名称(Null)', '(分录)单价', '(分录)数量', '(分录)原币金额', 
    '(分录)借方金额', '(分录)贷方金额', '(分录)结算方式#编码', '(分录)结算方式#名称(Null)', 
    '(分录)结算号', '(分录)业务单号', '(分录)失效分录行号'
]

if source_file and ratio_file:
    try:
        # 1. 智能解析TB费用表（升级模糊容错识别）
        src_excel = pd.ExcelFile(source_file)
        valid_src_sheet = src_excel.sheet_names[0]
        tb_header_idx = 0
        
        for sheet in src_excel.sheet_names:
            df_check = pd.read_excel(source_file, sheet_name=sheet, header=None)
            for idx, row in df_check.iterrows():
                row_strs = [str(v).strip() for v in row.values if pd.notna(v)]
                # 只要有一行里包含“待拆分”或“金额”相关核心字眼，即视为表头行
                if any("待拆分" in s or "待拆分金额" in s for s in row_strs):
                    valid_src_sheet = sheet
                    tb_header_idx = idx
                    break
                    
        df_source = pd.read_excel(source_file, sheet_name=valid_src_sheet, skiprows=tb_header_idx)
        df_source.columns = df_source.columns.astype(str).str.strip()
        
        # 智能锁定“待拆分金额”列名（容忍用户改名或带空格）
        amt_col = None
        for col in df_source.columns:
            if "待拆分" in col:
                amt_col = col
                break
        if not amt_col:
            st.error("❌ 在TB明细表中未找到包含‘待拆分金额’的列，请检查表格列名！")
            st.stop()
            
        # 2. 解析比例表双表头
        df_ratio_raw = pd.read_excel(ratio_file, header=None)
        ratio_header_idx = 0
        for idx, row in df_ratio_raw.iterrows():
            if "成本中心编号" in row.values:
                ratio_header_idx = idx
                break
        top_headers = df_ratio_raw.iloc[ratio_header_idx - 1].tolist() if ratio_header_idx > 0 else df_ratio_raw.iloc[0].tolist()
        text_headers = df_ratio_raw.iloc[ratio_header_idx].tolist()
        
        proj_text_to_code = {}
        fixed_names = ['成本中心编号', '成本中心名称', '大部门分类', '项目', '分摊逻辑', '分摊过渡部门', '分摊类型', '合计']
        for t_name, top_code in zip(text_headers, top_headers):
            if pd.notna(t_name) and pd.notna(top_code):
                t_str = str(t_name).strip()
                c_str = str(top_code).strip()
                if t_str not in fixed_names and not t_str.startswith('Unnamed:'):
                    if c_str.endswith('.0'):
                        c_str = c_str.split('.')[0].zfill(3)
                    else:
                        c_str = c_str.zfill(3)
                    proj_text_to_code[t_str] = c_str
        
        df_ratio = pd.read_excel(ratio_file, skiprows=ratio_header_idx)
        df_ratio.columns = df_ratio.columns.astype(str).str.strip()
        
        st.success(f"✅ 费用表表头锁定成功！识别到的金额列为: 【{amt_col}】")
        
        if st.button("🚀 开始全自动重分类并导出金蝶Excel"):
            project_cols = list(proj_text_to_code.keys())
            df_source['待拆分金额_numeric'] = pd.to_numeric(df_source[amt_col], errors='coerce')
            df_to_split = df_source[df_source['待拆分金额_numeric'].notna() & (df_source['待拆分金额_numeric'] != 0)]
            
            if df_to_split.empty:
                st.warning("⚠️ 没有发现有效数字数据行（或待拆分金额全部为0）。")
            
            base_info = {
                'FBillHead(GL_VOUCHER)': 1,
                'FAccountBookID': comp_info['book_id'],
                'FAccountBookID#Name': selected_company,
                'FDate': voucher_date,
                'FBUSDATE': voucher_date,
                'FYEAR': int(current_year),
                'FPERIOD': int(current_period),
                'FVOUCHERGROUPID': 'PRE001',
                'FVOUCHERGROUPID#Name': '记',
                'FVOUCHERGROUPNO': '1',
                'FATTACHMENTS': 0,
                'FISADJUSTVOUCHER': '否',
                'FACCBOOKORGID': comp_info['org_id'],
                'FACCBOOKORGID#Name': selected_company,
                'FCURRENCYID': chosen_currency['id'],
                'FCURRENCYID#Name': chosen_currency['name'],
                'FEXCHANGERATETYPE': 'HLTX01_SYS',
                'FEXCHANGERATETYPE#Name': '固定汇率',
                'FEXCHANGERATE': 1
            }
            
            new_rows = []
            entry_idx = 1
            
            for idx, row in df_to_split.iterrows():
                try:
                    sub_code = str(row['科目编码']).strip()
                    sub_name = str(row['科目名称']).strip()
                    dim_code = str(row['核算维度编码']).strip()
                    if sub_code == '科目编码' or sub_code == 'nan': continue
                        
                    orig_amt = float(row['待拆分金额_numeric'])
                    cc_code = None
                    parts = dim_code.split('/')
                    for p in parts:
                        p_clean = p.strip()
                        if p_clean in df_ratio['成本中心编号'].astype(str).str.strip().values:
                            cc_code = p_clean
                            break
                    if not cc_code and dim_code in df_ratio['成本中心编号'].astype(str).str.strip().values:
                        cc_code = dim_code
                    
                    if not cc_code: continue
                        
                    ratio_row_data = df_ratio[df_ratio['成本中心编号'].astype(str).str.strip() == cc_code].iloc[0]
                    valid_projects = []
                    for proj_text in project_cols:
                        val = ratio_row_data[proj_text]
                        try:
                            val_float = float(val)
                            if val_float > 0:
                                valid_projects.append({'proj_text': proj_text, 'proj_code': proj_text_to_code[proj_text], 'ratio': val_float})
                        except:
                            pass
                            
                    if not valid_projects: continue
                    df_valid_proj = pd.DataFrame(valid_projects).sort_values(by='ratio', ascending=False)
                    
                    # 1. 冲销行：借方填待拆分金额的负数
                    neg_row = [None] * len(tech_headers)
                    if entry_idx == 1:
                        for k, v in base_info.items():
                            if k in tech_headers: neg_row[tech_headers.index(k)] = v
                    
                    neg_row[tech_headers.index('FEntity')] = entry_idx
                    neg_row[tech_headers.index('FEXPLANATION')] = f"重分类-冲原公摊-{sub_name}"
                    neg_row[tech_headers.index('FACCOUNTID')] = sub_code
                    neg_row[tech_headers.index('FACCOUNTID#Name')] = sub_name
                    neg_row[tech_headers.index('FDEBIT')] = -orig_amt
                    neg_row[tech_headers.index('FCREDIT')] = None
                    neg_row[tech_headers.index('FAMOUNTFOR')] = None
                    
                    for field in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']:
                        neg_row[tech_headers.index(field)] = base_info[field]
                    
                    if 'FDetailID#FF100002' in tech_headers: neg_row[tech_headers.index('FDetailID#FF100002')] = '006'
                    if 'FDetailID#FFlex5' in tech_headers: neg_row[tech_headers.index('FDetailID#FFlex5')] = cc_code
                    
                    new_rows.append(neg_row)
                    entry_idx += 1
                    
                    # 2. 分配行：借方正数
                    allocated_sum = 0.0
                    for i, p_row in enumerate(df_valid_proj.itertuples()):
                        is_last = (i == len(df_valid_proj) - 1)
                        current_ratio = p_row.ratio
                        
                        if is_last:
                            amt = round(orig_amt - allocated_sum, 2)
                        else:
                            amt = round(orig_amt * current_ratio, 2)
                            allocated_sum += amt
                            
                        if amt == 0: continue
                        
                        pos_row = [None] * len(tech_headers)
                        pos_row[tech_headers.index('FEntity')] = entry_idx
                        pos_row[tech_headers.index('FEXPLANATION')] = f"重分类-项目分摊-{sub_name}"
                        pos_row[tech_headers.index('FACCOUNTID')] = sub_code
                        pos_row[tech_headers.index('FACCOUNTID#Name')] = sub_name
                        pos_row[tech_headers.index('FDEBIT')] = amt
                        pos_row[tech_headers.index('FCREDIT')] = None
                        pos_row[tech_headers.index('FAMOUNTFOR')] = None
                        
                        for field in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']:
                            pos_row[tech_headers.index(field)] = base_info[field]
                        
                        if 'FDetailID#FF100002' in tech_headers: pos_row[tech_headers.index('FDetailID#FF100002')] = p_row.proj_code
                        if 'FDetailID#FFlex5' in tech_headers: pos_row[tech_headers.index('FDetailID#FFlex5')] = cc_code
                        
                        new_rows.append(pos_row)
                        entry_idx += 1
                except:
                    pass
            
            if new_rows:
                final_df = pd.DataFrame([tech_headers, cn_headers] + new_rows)
                st.success(f"🎉 成功生成凭证！")
                st.dataframe(final_df.iloc[2:15])
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, header=False, sheet_name='凭证#单据头(FBillHead)')
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 点击下载金蝶引入Excel凭证",
                    data=processed_data,
                    file_name=f"金蝶云星空重分类凭证-{selected_company}-{chosen_currency['id']}-{voucher_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"处理发生意外错误: {e}")
else:
    st.info("💡 请在上方上传费用表与分摊比例表格。")

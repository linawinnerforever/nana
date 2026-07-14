# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="金蝶云星空-集团费重分类全自动工具", layout="wide")

st.title("📊 金蝶云星空 - 费用重分类集团全自动生成凭证工具")
st.markdown("""
上传费用明细与分摊比例，左侧直接选主体和期间，火箭一击即中！
""")

# ----------------------------------------------------
# 集团 18 大主体真实财务静态配置库
# ----------------------------------------------------
COMPANY_CONFIG = {
    "Crazy Maple Studio Inc": {"book_id": "100", "org_id": "002"},
    "CRAZY MAPLE  SERVICE  COMPANY": {"book_id": "101", "org_id": "003"},
    "Crazy Maple  Canada  Inc": {"book_id": "102", "org_id": "004"},
    "Maple House Inc": {"book_id": "103", "org_id": "005"},
    "CRAZY MAPLE  INTERACTIVE  HOLDING LTD": {"book_id": "104", "org_id": "006"},
    "SPICY MAPLE  LIMITED": {"book_id": "105", "org_id": "007"},
    "CRAZY MAPLE  STUDIO  HK LIMITED": {"book_id": "106", "org_id": "008"},
    "New Leaf  Publishing  Inc": {"book_id": "107", "org_id": "009"},
    "北京枫悦互动科技有限公司": {"book_id": "108", "org_id": "010"},
    "深圳枫叶互动科技有限公司": {"book_id": "109", "org_id": "011"},
    "杭州枫叶互动科技有限公司": {"book_id": "110", "org_id": "012"},
    "B25 LIMITED": {"book_id": "111", "org_id": "013"},
    "海南枫悦互动科技有限公司": {"book_id": "112", "org_id": "014"},
    "北京枫悦互动科技有限公司工会委员会": {"book_id": "113", "org_id": "015"},
    "深圳市星尘游戏科技有限公司": {"book_id": "114", "org_id": "016"},
    "ReelShort Japan Co., Ltd.": {"book_id": "115", "org_id": "017"},
    "SWEET MAPLE LIMITED": {"book_id": "116", "org_id": "018"},
    "深圳枫悦互动科技有限公司": {"book_id": "117", "org_id": "019"}
}

# ----------------------------------------------------
# 侧边栏：多主体与期间控制台
# ----------------------------------------------------
st.sidebar.header("🛠️ 集团财务控制台")

# 1. 动态选择公司主体
company_options = list(COMPANY_CONFIG.keys())
selected_company = st.sidebar.selectbox("请选择本次做账的公司主体", options=company_options)

comp_info = COMPANY_CONFIG[selected_company]
st.sidebar.success(f"📍 锁定主体：账簿({comp_info['book_id']}) | 组织({comp_info['org_id']})")

# 2. 动态选择期间
current_year = st.sidebar.number_input("会计年度 (FYEAR)", min_value=2020, max_value=2035, value=2026)
current_period = st.sidebar.slider("会计期间 (FPERIOD)", min_value=1, max_value=12, value=6)

# 自动计算当月最后一天
last_day = calendar.monthrange(current_year, current_period)[1]
voucher_date = f"{current_year}-{str(current_period).zfill(2)}-{str(last_day).zfill(2)}"
st.sidebar.info(f"📅 凭证自动生成日期：{voucher_date}")

# ----------------------------------------------------
# 主界面：极简两个上传框
# ----------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 上传费用明细表 (TB)")
    source_file = st.file_uploader("支持 .xlsx 格式，需含：科目编码、科目名称、核算维度编码、核算维度名称、待拆分金额", type=["xlsx"], key="src")

with col2:
    st.subheader("2. 上传部门项目分摊比例表")
    ratio_file = st.file_uploader("支持 .xlsx 格式，需含：第一行数字编码，第二行文字名称的二维比例矩阵", type=["xlsx"], key="ratio")

# 金蝶云星空底层固定的 43 列标准表头架构
tech_headers = [
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
    'FDetailID#FF100005#Name', 'FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 
    'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE', 'FUnitId', 'FUnitId#Name', 'FPrice', 'FQty', 
    'FAMOUNTFOR', 'FDEBIT', 'FCREDIT', 'FSettleTypeID', 'FSettleTypeID#Name', 'FSETTLENO', 
    'FBUSNO', 'FEXPORTENTRYID'
]

cn_headers = [
    '*单据头(序号)', '*(单据头)账簿#编码', '(单据头)账簿#名称', '*(单据头)日期', '(单据头)业务日期', 
    '(单据头)会计年度', '(单据头)期间', '*(单据头)凭证字#编码', '(单据头)凭证字#名称', 
    '*(单据头)凭证号', '(单据头)附件数', '(单据头)是否调整凭证', '*(单据头)核算组织#编码', 
    '(单据头)核算组织#名称', '(单据头)来源系统', '(单据头)来源系统#名称', '(单据头)引入版本', 
    '*分录(序号)', 'FEntity', '摘要', '科目控制#编码', '科目控制#名称', '往来单位#编码', 
    '往来单位#名称', '项目#编码', '项目#名称', '第二核算维度#编码', '第二核算维度#名称', 
    '第三核算维度#编码', '第三核算维度#名称', '费用类别#编码', '费用类别#名称', '备用#编码', 
    '备用#名称', '第四核算维度#编码', '第四核算维度#名称', '第五核算维度#编码', '第五核算维度#名称', 
    '第六核算维度#编码', '第六核算维度#名称', '第七核算维度#编码', '第七核算维度#名称', 
    '第八核算维度#编码', '第八核算维度#名称', '第九核算维度#编码', '第九核算维度#名称', 
    '个人#编码', '个人#名称', '部门#编码', '部门#名称', '第十核算维度#编码', '第十核算维度#名称', 
    '收支项目#编码', '收支项目#名称', '资金计划项目#编码', '资金计划项目#名称', '原币#编码', 
    '原币#名称', '汇率类型#编码', '汇率类型#名称', '汇率', '计量单位#编码', '计量单位#名称', 
    '单价', '数量', '原币金额', '借方金额', '贷方金额', '结算方式#编码', '结算方式#名称', 
    '结算号', '业务单号', '引入分录失效标志'
]

if source_file and ratio_file:
    try:
        # 1. 解析TB费用表
        src_excel = pd.ExcelFile(source_file)
        valid_src_sheet = src_excel.sheet_names[0]
        tb_header_idx = 0
        for sheet in src_excel.sheet_names:
            df_check = pd.read_excel(source_file, sheet_name=sheet, header=None)
            for idx, row in df_check.iterrows():
                if "待拆分金额" in row.values:
                    valid_src_sheet = sheet
                    tb_header_idx = idx
                    break
        df_source = pd.read_excel(source_file, sheet_name=valid_src_sheet, skiprows=tb_header_idx)
        df_source.columns = df_source.columns.astype(str).str.strip()
        
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
        
        st.success(f"✅ 文件加载就绪！当前做账主体为: 【{selected_company}】")
        
        if st.button("🚀 开始全自动重分类并导出金蝶Excel"):
            project_cols = list(proj_text_to_code.keys())
            df_source['待拆分金额_numeric'] = pd.to_numeric(df_source['待拆分金额'], errors='coerce')
            df_to_split = df_source[df_source['待拆分金额_numeric'].notna() & (df_source['待拆分金额_numeric'] != 0)]
            
            if df_to_split.empty:
                st.warning("⚠️ 没有发现金额不为0的有效数字数据行。")
            
            # 自动注入匹配的正确金蝶编码
            base_info = {
                'FAccountBookID': comp_info['book_id'],
                'FAccountBookID#Name': selected_company,
                'FACCBOOKORGID': comp_info['org_id'],
                'FACCBOOKORGID#Name': selected_company,
                'FDate': voucher_date,
                'FBUSDATE': voucher_date,
                'FYEAR': int(current_year),
                'FPERIOD': int(current_period),
                'FVOUCHERGROUPID': '記',
                'FVOUCHERGROUPID#Name': '记账凭证',
                'FVOUCHERGROUPNO': '1',
                'FATTACHMENTS': 0,
                'FISADJUSTVOUCHER': '否',
                'FCURRENCYID': 'PRE001',
                'FCURRENCYID#Name': '人民币',
                'FEXCHANGERATETYPE': 'HLTX01_SYS',
                'FEXCHANGERATETYPE#Name': '固定汇率',
                'FEXCHANGERATE': 1
            }
            
            new_rows = []
            entry_idx = 1
            split_idx = tech_headers.index('*Split*1')
            
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
                    
                    if not cc_code:
                        continue
                        
                    ratio_row_data = df_ratio[df_ratio['成本中心编号'].astype(str).str.strip() == cc_code].iloc[0]
                    
                    valid_projects = []
                    for proj_text in project_cols:
                        val = ratio_row_data[proj_text]
                        try:
                            val_float = float(val)
                            if val_float > 0:
                                valid_projects.append({
                                    'proj_text': proj_text, 
                                    'proj_code': proj_text_to_code[proj_text], 
                                    'ratio': val_float
                                })
                        except:
                            pass
                            
                    if not valid_projects: continue
                    
                    df_valid_proj = pd.DataFrame(valid_projects).sort_values(by='ratio', ascending=False)
                    
                    # 1. 借方负数行 (首行填充 A-R)
                    neg_row = [None] * len(tech_headers)
                    for k, v in base_info.items():
                        if k in tech_headers: neg_row[tech_headers.index(k)] = v
                    
                    neg_row[tech_headers.index('FEntity')] = entry_idx
                    neg_row[tech_headers.index('FEXPLANATION')] = f"重分类-冲原公摊-{sub_name}"
                    neg_row[tech_headers.index('FACCOUNTID')] = sub_code
                    neg_row[tech_headers.index('FACCOUNTID#Name')] = sub_name
                    neg_row[tech_headers.index('FDEBIT')] = -orig_amt
                    neg_row[tech_headers.index('FAMOUNTFOR')] = -orig_amt
                    
                    if 'FDetailID#FF100002' in tech_headers: neg_row[tech_headers.index('FDetailID#FF100002')] = '006'
                    if 'FDetailID#FFlex5' in tech_headers: neg_row[tech_headers.index('FDetailID#FFlex5')] = cc_code
                    
                    new_rows.append(neg_row)
                    entry_idx += 1
                    
                    # 2. 借方正数行 (后续行 A-R 完美留空)
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
                        for k, v in base_info.items():
                            if k in tech_headers:
                                k_idx = tech_headers.index(k)
                                if k_idx >= split_idx:
                                    pos_row[k_idx] = v
                        
                        pos_row[tech_headers.index('FVOUCHERGROUPID')] = base_info['FVOUCHERGROUPID']
                        pos_row[tech_headers.index('FVOUCHERGROUPNO')] = base_info['FVOUCHERGROUPNO']
                        
                        pos_row[tech_headers.index('FEntity')] = entry_idx
                        pos_row[tech_headers.index('FEXPLANATION')] = f"重分类-项目分摊-{sub_name}"
                        pos_row[tech_headers.index('FACCOUNTID')] = sub_code
                        pos_row[tech_headers.index('FACCOUNTID#Name')] = sub_name
                        pos_row[tech_headers.index('FDEBIT')] = amt
                        pos_row[tech_headers.index('FAMOUNTFOR')] = amt
                        
                        if 'FDetailID#FF100002' in tech_headers: pos_row[tech_headers.index('FDetailID#FF100002')] = p_row.proj_code
                        if 'FDetailID#FFlex5' in tech_headers: pos_row[tech_headers.index('FDetailID#FFlex5')] = cc_code
                        
                        new_rows.append(pos_row)
                        entry_idx += 1
                        
                except:
                    pass
            
            if new_rows:
                final_df = pd.DataFrame([tech_headers, cn_headers] + new_rows)
                st.success(f"🎉 成功生成 【{selected_company}】 凭证！")
                st.dataframe(final_df.iloc[2:15])
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 点击下载金蝶引入Excel凭证",
                    data=processed_data,
                    file_name=f"金蝶云星空重分类凭证-{selected_company}-{voucher_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"处理发生意外错误: {e}")
else:
    st.info("💡 请在上方上传费用表与分摊比例表格。")

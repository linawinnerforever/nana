# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="独立版-投放费用自动金蝶凭证工具", layout="wide")

st.title("📈 投放费用全自动金蝶凭证生成器 (独立核算版)")
st.markdown("---")

# ----------------------------------------------------
# 🔐 独立链接的私人专属口令锁
# ----------------------------------------------------
st.sidebar.markdown("### 🔑 私人安全网闸")
access_token = st.sidebar.text_input("请输入专属理账口令：", type="password")

if access_token != "lina":
    st.warning("🔒 请输入正确的私人授权口令以解锁专属理账面板。")
    st.stop()

st.sidebar.success("🔓 投放费用独立理账舱已安全解锁。")

# ----------------------------------------------------
# 👑 2026最新产品核算项目编码与三大主体（CM/MH/NL）对齐真理库
# ----------------------------------------------------
PRODUCT_MAIN_MAP = {
    "CHAPTERS": {"company_name": "Crazy Maple Studio Inc", "book_id": "002", "org_id": "100", "project_code": "002", "main_body": "CM"},
    "MERGE": {"company_name": "Crazy Maple Studio Inc", "book_id": "002", "org_id": "100", "project_code": "004", "main_body": "CM"},
    "KISS": {"company_name": "Maple House Inc", "book_id": "005", "org_id": "103", "project_code": "003", "main_body": "MH"},
    "REELSHORT": {"company_name": "New Leaf Publishing Inc", "book_id": "009", "org_id": "107", "project_code": "005", "main_body": "NL"},
    "MAXDRAMA": {"company_name": "New Leaf Publishing Inc", "book_id": "009", "org_id": "107", "project_code": "001", "main_body": "NL"},
    "RS N": {"company_name": "New Leaf Publishing Inc", "book_id": "009", "org_id": "107", "project_code": "006", "main_body": "NL"},
    "RSNOVEL": {"company_name": "New Leaf Publishing Inc", "book_id": "009", "org_id": "107", "project_code": "006", "main_body": "NL"}
}

# 金蝶引入 73 列对照表头
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
    '*(分录)币别#编码', '(分录)币别#名称', '*(分录)汇率类型#编码', '(分录)汇率类型#名称', '(分录)汇率', 
    '(分录)计量单位#编码', '(分录)计量单位#名称(Null)', '(分录)单价', '(分录)数量', '(分录)原币金额', 
    '(分录)借方金额', '(分录)贷方金额', '(分录)结算方式#编码', '(分录)结算方式#名称(Null)', 
    '(分录)结算号', '(分录)业务编号', '(分录)现金流量#分录ID'
]

# ----------------------------------------------------
# 期间控制台
# ----------------------------------------------------
st.sidebar.header("📅 会计期间参数")
current_year = st.sidebar.number_input("会计年度 (FYEAR)", min_value=2020, max_value=2035, value=2026)
current_period = st.sidebar.slider("会计期间 (FPERIOD)", min_value=1, max_value=12, value=6)
last_day = calendar.monthrange(current_year, current_period)[1]
voucher_date = f"{current_year}-{str(current_period).zfill(2)}-{str(last_day).zfill(2)}"

# ----------------------------------------------------
# 业务文件多表上传处理区
# ----------------------------------------------------
st.markdown("### 📤 第一步：上传数据源资产")
col_up1, col_up2 = st.columns(2)
with col_up1:
    ad_files = st.file_uploader("1. 请选择上传您的产品买量计提表 Excel（支持多选）", type=["xlsx"], accept_multiple_files=True)
with col_up2:
    mp_file = st.file_uploader("2. 请上传您的金蝶供应商映射真理表 (投放费用MP.xlsx)", type=["xlsx"])

if ad_files and mp_file:
    try:
        # 🌟 安全嗅探加载供应商表
        try:
            df_vendor = pd.read_excel(mp_file, sheet_name=0)
        except:
            df_vendor = pd.read_excel(mp_file, sheet_name=0, engine='openpyxl')
            
        df_vendor.columns = df_vendor.columns.astype(str).str.strip()
        
        vendor_code_map = {}
        vendor_name_map = {}
        for idx, row in df_vendor.iterrows():
            chan = str(row.get("渠道", "")).strip().upper()
            v_code = str(row.get("编码", "")).strip()
            v_name = str(row.get("金蝶", "")).strip()
            if chan and chan != "NAN" and v_code and v_code != "NAN":
                vendor_code_map[chan] = v_code
                vendor_name_map[chan] = v_name
                
        st.success("📊 供应商级联字典配置加载就绪！")
        
        st.markdown("### ⚡ 第二步：执行一键对账汇总")
        if st.button("🚀 启动全自动洗流并输出 CM & MH 专属做账资产"):
            master_raw_list = []
            has_error_flag = False
            missing_vendors_set = set()
            voucher_rows = []
            ent_id = 1
            
            # 循环清洗所有业务表格
            for file in ad_files:
                # 🌟【核心修复点】：采用智能动态安全流引擎读取，100%避开旧版 xlrd/openpyxl 格式歧义
                try:
                    df_sheet_raw = pd.read_excel(file, sheet_name=0, header=None)
                except:
                    try:
                        df_sheet_raw = pd.read_excel(file, sheet_name=0, header=None, engine='openpyxl')
                    except:
                        # 兜底旧格式降级兼容
                        df_sheet_raw = pd.read_excel(file, sheet_name=0, header=None, engine='xlrd')
                
                header_row_idx = 0
                for r_idx, r_val in df_sheet_raw.iterrows():
                    r_strs = [str(v).strip() for v in r_val.values if pd.notna(v)]
                    if "投放渠道" in r_strs or "项目" in r_strs or "买量产品" in r_strs:
                        header_row_idx = r_idx
                        break
                
                # 重新用锁定到的真实行首进行切片读取
                try:
                    df_clean = pd.read_excel(file, sheet_name=0, skiprows=header_row_idx)
                except:
                    df_clean = pd.read_excel(file, sheet_name=0, skiprows=header_row_idx, engine='openpyxl')
                    
                df_clean.columns = df_clean.columns.astype(str).str.strip()
                
                if "项目" in df_clean.columns:
                    df_clean.rename(columns={"项目": "买量产品"}, inplace=True)
                if "投放渠道" not in df_clean.columns:
                    continue
                    
                for idx, row in df_clean.iterrows():
                    p_name_raw = str(row.get("买量产品", "")).strip()
                    p_name = p_name_raw.upper()
                    channel = str(row.get("投放渠道", "")).strip()
                    partner = str(row.get("开户方", "")).strip()
                    
                    if p_name == "" or p_name == "NAN" or "合计" in p_name_raw or "求和" in p_name_raw:
                        continue
                    if channel == "" or channel == "NAN" or "合计" in channel:
                        continue
                        
                    # 划分主体真理库过滤
                    if p_name not in PRODUCT_MAIN_MAP:
                        continue
                    prod_meta = PRODUCT_MAIN_MAP[p_name]
                    
                    # 金额清洗
                    spent_raw = row.get("spent", 0)
                    spent_str = str(spent_raw).replace(",", "").strip()
                    # 兼容可能存在的业务表格特殊符号减号
                    spent_str = spent_str.replace("−", "-").replace("—", "-")
                    
                    try:
                        spent_val = float(spent_str) if spent_str != "" else 0.0
                    except:
                        spent_val = 0.0
                        
                    if spent_val == 0:
                        continue
                        
                    # 供应商-渠道动态判定
                    if partner and partner != "nan" and partner != "(空白)":
                        supplier_channel = partner
                    else:
                        supplier_channel = channel
                        
                    # 金蝶信息级联匹配
                    match_key = supplier_channel.strip().upper()
                    final_v_code = vendor_code_map.get(match_key, "")
                    final_v_kingdee = vendor_name_map.get(match_key, "")
                    
                    if not final_v_code:
                        final_v_code = "🚨 未匹配到，请检查供应商表"
                        final_v_kingdee = "🚨 未匹配到，请检查供应商表"
                        has_error_flag = True
                        missing_vendors_set.add(supplier_channel)
                        
                    # 摘要截断
                    if partner and partner != "nan" and partner != "(空白)":
                        explanation_str = f"计提{current_year}年{current_period}月推广费用-Jun.{current_year} advertising and marketing cost accrual-{channel}-{partner}"
                        partner_label = partner
                    else:
                        explanation_str = f"计提{current_year}年{current_period}月推广费用-Jun.{current_year} advertising and marketing cost accrual-{channel}"
                        partner_label = "(空白)"
                        
                    record = {
                        "买量产品": prod_meta["product_label"],
                        "投放渠道": channel,
                        "开户方": partner_label,
                        "求和项:spent": spent_val,
                        "主体": prod_meta["main_body"],
                        "核算项目": prod_meta["project_code"],
                        "供应商-渠道": supplier_channel,
                        "供应商-金蝶": final_v_kingdee,
                        "供应商编码": final_v_code,
                        "摘要": explanation_str
                    }
                    master_raw_list.append(record)
                    
                    # 组装金蝶凭证行
                    base_info = {
                        'FBillHead(GL_VOUCHER)': 1, 'FAccountBookID': prod_meta["book_id"], 'FAccountBookID#Name': prod_meta["company_name"],
                        'FDate': voucher_date, 'FBUSDATE': voucher_date, 'FYEAR': int(current_year), 'FPERIOD': int(current_period),
                        'FVOUCHERGROUPID': 'PRE001', 'FVOUCHERGROUPID#Name': '记', 'FVOUCHERGROUPNO': str(ent_id), 'FATTACHMENTS': 0, 'FISADJUSTVOUCHER': '否',
                        'FACCBOOKORGID': prod_meta["org_id"], 'FACCBOOKORGID#Name': prod_meta["company_name"],
                        'FCURRENCYID': 'PRE007', 'FCURRENCYID#Name': '美元', 'FEXCHANGERATETYPE': 'HLTX01_SYS', 'FEXCHANGERATETYPE#Name': '固定汇率', 'FEXCHANGERATE': 1
                    }
                    # 借方
                    d_row = [None] * len(tech_headers)
                    for k, v in base_info.items(): d_row[tech_headers.index(k)] = v
                    d_row[tech_headers.index('FEntity')] = 1
                    d_row[tech_headers.index('FEXPLANATION')] = explanation_str
                    d_row[tech_headers.index('FACCOUNTID')] = "6601.03.01"
                    d_row[tech_headers.index('FACCOUNTID#Name')] = "销售费用_广告宣传费_买量投放"
                    d_row[tech_headers.index('FDEBIT')] = spent_val
                    if 'FDetailID#FF100002' in tech_headers: d_row[tech_headers.index('FDetailID#FF100002')] = prod_meta["project_code"]
                    if 'FDetailID#FFlex5' in tech_headers: d_row[tech_headers.index('FDetailID#FFlex5')] = "7000"
                    voucher_rows.append(d_row)
                    # 贷方
                    c_row = [None] * len(tech_headers)
                    c_row[tech_headers.index('FEntity')] = 2
                    c_row[tech_headers.index('FEXPLANATION')] = explanation_str
                    c_row[tech_headers.index('FACCOUNTID')] = "2202.02"
                    c_row[tech_headers.index('FACCOUNTID#Name')] = "应付账款_费用服务费"
                    c_row[tech_headers.index('FCREDIT')] = spent_val
                    if "🚨" not in final_v_code and 'FDetailID#FF100005' in tech_headers:
                        c_row[tech_headers.index('FDetailID#FF100005')] = final_v_code
                    for f in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']: c_row[tech_headers.index(f)] = base_info[f]
                    voucher_rows.append(c_row)
                    ent_id += 1

            if not master_raw_list:
                st.error("❌ 未能从这些 Excel 表格中清洗出任何属于指定主体的买量流数据。")
                st.stop()
                
            df_ledger = pd.DataFrame(master_raw_list)
            df_voucher = pd.DataFrame([tech_headers, cn_headers] + voucher_rows)
            df_pivot = df_ledger.groupby(["主体", "买量产品", "供应商-金蝶"])["求和项:spent"].sum().reset_index()
            df_pivot.rename(columns={"求和项:spent": "总金额(USD)"}, inplace=True)
            
            if has_error_flag:
                st.error(f"🚨 【注意】检测到有 {len(missing_vendors_set)} 个新增的“供应商-渠道”在映射表里不存在：")
                st.warning(f"缺失的渠道名为：{list(missing_vendors_set)}")
            else:
                st.success("✨ 所有主体服务商无缝完美匹配！")
                
            st.markdown("#### 📊 当期买量消耗分类核对看板")
            st.dataframe(df_pivot, use_container_width=True)
            
            output_stream = io.BytesIO()
            with pd.ExcelWriter(output_stream, engine='xlsxwriter') as writer:
                df_ledger.to_excel(writer, index=False, sheet_name='整理后入账底稿(todo)')
                df_voucher.to_excel(writer, index=False, header=False, sheet_name='金蝶导入凭证')
                df_pivot.to_excel(writer, index=False, sheet_name='透视分类看板')
                
            st.download_button(
                label="📥 导出更新后的完整投放理账大资产包",
                data=output_stream.getvalue(),
                file_name=f"投放费用整理入账底稿_{voucher_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"处理文件出现技术意外，详情: {e}")

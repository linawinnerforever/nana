# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import calendar
from pypdf import PdfReader

# 1. 页面主配置
st.set_page_config(page_title="金蝶云星空-集团财务全自动工具箱", layout="wide")

st.title("📊 金蝶云星空 - 集团财务全自动高阶工具箱")

# ----------------------------------------------------
# 共享配置真理库（账簿与组织编码100%对齐最新版）
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

EMPLOYEE_DEPT_MAP = {
    "Yapeng Nan": {"dept_code": "5000", "emp_code": "0002"},
    "Yi Jia": {"dept_code": "5000", "emp_code": "0001"},
    "Kruti": {"dept_code": "5000", "emp_code": "0003"},
    "Tania": {"dept_code": "2005", "emp_code": "0064"}
}

MERCHANT_RULES = [
    {"keyword": "DEEPL", "project": "订阅类支出", "acct_code": "6602.04", "acct_name": "管理费用_办公费"},
    {"keyword": "ANTHROPIC", "project": "软件使用费-ANTHROPIC", "acct_code": "6401.21", "acct_name": "主营业务成本_软件服务费", "vendor": "VEN02027"},
    {"keyword": "Google CLOUD", "project": "软件使用费-Google Cloud", "acct_code": "6401.21", "acct_name": "主营业务成本_软件服务费", "vendor": "VEN00057"},
    {"keyword": "GOOGLE *ReelShort", "project": "行政办公类支出", "acct_code": "6602.04", "acct_name": "管理费用_办公费"},
    {"keyword": "GOOGLE *SVCSCRAZYMAPLES", "project": "订阅类支出", "acct_code": "6602.04", "acct_name": "管理费用_办公费"},
    {"keyword": "ADOBE", "project": "订阅类支出", "acct_code": "6602.04", "acct_name": "管理费用_办公费"},
    {"keyword": "ZOOM.COM", "project": "订阅类支出", "acct_code": "6602.04", "acct_name": "管理费用_办公费"},
    {"keyword": "UBER *ONE", "project": "主营成本-员工福利", "acct_code": "6602.01.02.02", "acct_name": "管理费用_职工薪酬_Employee_职工福利"},
    {"keyword": "OLIVE GARDEN", "project": "主营成本-员工福利", "acct_code": "6602.01.02.02", "acct_name": "管理费用_职工薪酬_Employee_职工福利"},
    {"keyword": "INTUIT", "project": "软件使用费-Office expense", "acct_code": "6602.04", "acct_name": "管理费用_办公费"},
    {"keyword": "OPENAI", "project": "软件使用费-OPENAI", "acct_code": "6401.21", "acct_name": "主营业务成本_软件服务费"},
    {"keyword": "COMPASS", "project": "行政办公类支出", "acct_code": "6602.04", "acct_name": "管理费用_办公费"}
]

# 🌟 镜像对齐的 73 列终极金蝶中英文对照表头
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
# 🔐 侧边栏：【密码锁分流控制器】口令：lina
# ----------------------------------------------------
st.sidebar.markdown("### 🚀 财务主控制台")
access_token = st.sidebar.text_input("请输入私人授权口令：", type="password")

if access_token == "lina":
    main_mode = st.sidebar.radio(
        "请选择您要执行的财务模块：",
        options=["📊 集团费用重分类工具", "💳 信用卡对账单智能理账(自用)"]
    )
else:
    main_mode = "📊 集团费用重分类工具"

# 侧边栏公用期间控制项
st.sidebar.header("🛠️ 期间及做账主体控制")
selected_company = st.sidebar.selectbox("请选择本次做账的公司主体", options=list(COMPANY_CONFIG.keys()))
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


# ====================================================
# 📦 密封舱一：💳 信用卡对账单智能理账（你专属自用）
# ====================================================
def run_credit_card_tool():
    st.subheader("💳 信用卡对账单（PDF）智能大闭环理账面板")
    
    def generate_voucher_dataframe(df_pivot_data, ali_amount, ali_debit, ali_credit):
        base_v_info = {
            'FBillHead(GL_VOUCHER)': 1, 'FAccountBookID': comp_info['book_id'], 'FAccountBookID#Name': selected_company,
            'FDate': voucher_date, 'FBUSDATE': voucher_date, 'FYEAR': int(current_year), 'FPERIOD': int(current_period),
            'FVOUCHERGROUPID': 'PRE001', 'FVOUCHERGROUPID#Name': '记', 'FVOUCHERGROUPNO': '1', 'FATTACHMENTS': 0, 'FISADJUSTVOUCHER': '否',
            'FACCBOOKORGID': comp_info['org_id'], 'FACCBOOKORGID#Name': selected_company,
            'FCURRENCYID': 'PRE007', 'FCURRENCYID#Name': '美元', 'FEXCHANGERATETYPE': 'HLTX01_SYS', 'FEXCHANGERATETYPE#Name': '固定汇率', 'FEXCHANGERATE': 1
        }
        voucher_rows = []
        ent_id = 1
        for p_row in df_pivot_data.itertuples():
            emp_info = EMPLOYEE_DEPT_MAP.get(p_row.持卡人, {"dept_code": "5000", "emp_code": ""})
            v_row = [None] * len(tech_headers)
            if ent_id == 1:
                for k, v in base_v_info.items(): v_row[tech_headers.index(k)] = v
            v_row[tech_headers.index('FEntity')] = ent_id
            v_row[tech_headers.index('FEXPLANATION')] = f"计提办公费-Office expense accrual-{p_row.持卡人}-{p_row.项目}"
            v_row[tech_headers.index('FACCOUNTID')] = str(p_row.科目编码).strip()
            v_row[tech_headers.index('FACCOUNTID#Name')] = p_row.科目名称
            v_row[tech_headers.index('FDEBIT')] = p_row.金额
            if 'FDetailID#FF100002' in tech_headers: v_row[tech_headers.index('FDetailID#FF100002')] = '001'
            if 'FDetailID#FFlex6' in tech_headers: v_row[tech_headers.index('FDetailID#FFlex6')] = emp_info["dept_code"]
            if 'FDetailID#FFlex4' in tech_headers: v_row[tech_headers.index('FDetailID#FFlex4')] = emp_info["emp_code"]
            if hasattr(p_row, '供应商') and p_row.供应商 and 'FDetailID#FF100005' in tech_headers: v_row[tech_headers.index('FDetailID#FF100005')] = p_row.供应商
            for f in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']: v_row[tech_headers.index(f)] = base_v_info[f]
            voucher_rows.append(v_row)
            ent_id += 1
            
            c_row = [None] * len(tech_headers)
            c_row[tech_headers.index('FEntity')] = ent_id
            c_row[tech_headers.index('FEXPLANATION')] = f"计提办公费-Office expense accrual-{p_row.持卡人}"
            c_row[tech_headers.index('FACCOUNTID')] = "2241.01"
            c_row[tech_headers.index('FCREDIT')] = p_row.金额
            if 'FDetailID#FFlex4' in tech_headers: c_row[tech_headers.index('FDetailID#FFlex4')] = emp_info["emp_code"]
            for f in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']: c_row[tech_headers.index(f)] = base_v_info[f]
            voucher_rows.append(c_row)
            ent_id += 1
        if ali_amount > 0:
            ali_d = [None] * len(tech_headers)
            ali_d[tech_headers.index('FEntity')] = ent_id
            ali_d[tech_headers.index('FEXPLANATION')] = f"支付阿里云服务费-Payment for Alibaba Cloud"
            ali_d[tech_headers.index('FACCOUNTID')] = ali_debit
            ali_d[tech_headers.index('FDEBIT')] = ali_amount
            if 'FDetailID#FF100005' in tech_headers: ali_d[tech_headers.index('FDetailID#FF100005')] = "VEN00057"
            for f in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']: ali_d[tech_headers.index(f)] = base_v_info[f]
            voucher_rows.append(ali_d)
            ent_id += 1
            
            ali_c = [None] * len(tech_headers)
            ali_c[tech_headers.index('FEntity')] = ent_id
            ali_c[tech_headers.index('FEXPLANATION')] = f"支付阿里云服务费-Payment for Alibaba Cloud"
            ali_c[tech_headers.index('FACCOUNTID')] = ali_credit
            ali_c[tech_headers.index('FCREDIT')] = ali_amount
            if 'FDetailID#FF100005' in tech_headers: ali_c[tech_headers.index('FDetailID#FF100005')] = "VEN00057"
            for f in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']: ali_c[tech_headers.index(f)] = base_v_info[f]
            voucher_rows.append(ali_c)
            ent_id += 1
        return pd.DataFrame([tech_headers, cn_headers] + voucher_rows)

    tab1, tab2 = st.tabs(["🆕 模式 1：从PDF账单一键生成大闭环", "✏️ 模式 2：上传修正底稿生成最终凭证"])
    st.markdown("##### ⚙️ 阿里云专项计提对冲参数")
    ali_col1, ali_col2, ali_col3 = st.columns(3)
    with ali_col1: ali_amt = st.number_input("阿里云本月金额 (USD)", min_value=0.0, value=1058598.35, key="ali_amt_sub")
    with ali_col2: ali_acct_debit = st.text_input("借方暂估科目", value="2202.01", key="ali_d_sub")
    with ali_col3: ali_acct_credit = st.text_input("贷方对冲科目", value="2241.03", key="ali_c_sub")

    with tab1:
        pdf_files = st.file_uploader("请在此投递信用卡对账单 PDF", type=["pdf"], accept_multiple_files=True, key="p_up")
        if pdf_files and st.button("🚀 启动全自动解析闭环逻辑"):
            extracted_tx = []
            for pdf_file in pdf_files:
                reader = PdfReader(pdf_file)
                current_person = "Unknown"
                for page in reader.pages:
                    text = page.extract_text()
                    if not text: continue
                    for line in text.split('\n'):
                        if "XXXX-XXXX-XXXX-" in line or "Account Number:" in line:
                            for name in EMPLOYEE_DEPT_MAP.keys():
                                if name.upper() in line.upper() or name.split()[-1].upper() in line.upper():
                                    current_person = name
                                    break
                        match = re.search(r'^(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(.+?)\s+([\d,]+\.\d{2})', line)
                        if match:
                            tx_date, post_date, desc, charge = match.groups()
                            charge_val = float(charge.replace(',', ''))
                            if "PAYMENT" in desc.upper() or "CREDIT" in desc.upper(): continue
                            extracted_tx.append({"期间": f"{current_year}-{str(current_period).zfill(2)}", "持卡人": current_person, "交易日期": tx_date, "原始商户描述": desc.strip(), "金额": charge_val})
            df_tx = pd.DataFrame(extracted_tx)
            
            # 💡 财务大脑规则库咬合映射（已修复解包报错）
            def route_accounting(row):
                desc = str(row['原始商户描述']).upper()
                project = "未分类支出(请在底稿修改)"
                acct_code = "6602.04"
                acct_name = "管理费用_办公费"
                vendor = ""
                for rule in MERCHANT_RULES:
                    if rule["keyword"].upper() in desc:
                        project = rule["project"]
                        acct_code = rule["acct_code"]
                        acct_name = rule["acct_name"]
                        vendor = rule.get("vendor", "")
                        break
                return project, acct_code, acct_name, vendor
                
            # 执行稳定咬合解析与列展开
            df_tx[['项目', '科目编码', '科目名称', '供应商']] = df_tx.apply(route_accounting, axis=1, result_type='expand')
            
            df_pivot = df_tx.groupby(['持卡人', '项目', '科目编码', '科目名称'])['金额'].sum().reset_index()
            df_voucher = generate_voucher_dataframe(df_pivot, ali_amt, ali_acct_debit, ali_acct_credit)
            
            output_blob = io.BytesIO()
            with pd.ExcelWriter(output_blob, engine='xlsxwriter') as writer:
                df_voucher.to_excel(writer, index=False, header=False, sheet_name='凭证#单据头(FBillHead)')
                df_pivot.to_excel(writer, index=False, sheet_name='透视看板')
                df_tx.to_excel(writer, index=False, sheet_name='信用卡拆分明细')
            final_data = output_blob.getvalue()
            st.success("🎉 一键闭环 Excel 生成完毕！金蝶导入第一页，账目对账和手工更正看后两页！")
            st.download_button(label="📥 点击下载【全功能大闭环Excel】", data=final_data, file_name=f"信用卡全功能闭环理账表-{selected_company}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        edited_file = st.file_uploader("请在此投递经过您手工修正未分类账目后的底稿 Excel", type=["xlsx"], key="ex_up")
        if edited_file and st.button("⚙️ 重新根据手工底稿转化最终金蝶凭证"):
            try:
                df_user_pivot = pd.read_excel(edited_file, sheet_name="透视看板")
                df_new_voucher = generate_voucher_dataframe(df_user_pivot, ali_amt, ali_acct_debit, ali_acct_credit)
                output_pure = io.BytesIO()
                with pd.ExcelWriter(output_pure, engine='xlsxwriter') as writer: df_new_voucher.to_excel(writer, index=False, header=False, sheet_name='凭证#单据头(FBillHead)')
                st.success("✨ 修正成功！已基于您的手工修正出具最终金蝶单页凭证。")
                st.download_button(label="📥 下载最终修正版金蝶引入凭证", data=output_pure.getvalue(), file_name=f"金蝶直接引入凭证(修正后)-{selected_company}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e: st.error(f"处理失败，请确保底稿包含‘透视看板’标签。错误详情: {e}")

# ====================================================
# 📊 密封舱二：老功能（老少咸宜的集团费用重分类工具）
# ====================================================
def run_reclassification_tool():
    st.subheader("📊 费用重分类集团全自动凭证生成板块")
    source_file = st.file_uploader("1. 上传费用明细表 (TB)", type=["xlsx"], key="src_old")
    ratio_file = st.file_uploader("2. 上传部门项目分摊比例表", type=["xlsx"], key="ratio_old")

    if source_file and ratio_file:
        try:
            src_excel = pd.ExcelFile(source_file)
            valid_src_sheet = src_excel.sheet_names[0]
            tb_header_idx = 0
            for sheet in src_excel.sheet_names:
                df_check = pd.read_excel(source_file, sheet_name=sheet, header=None)
                for idx, row in df_check.iterrows():
                    row_strs = [str(v).strip() for v in row.values if pd.notna(v)]
                    if any("待拆分" in s or "待拆分金额" in s for s in row_strs):
                        valid_src_sheet = sheet; tb_header_idx = idx; break
            df_source = pd.read_excel(source_file, sheet_name=valid_src_sheet, skiprows=tb_header_idx)
            df_source.columns = df_source.columns.astype(str).str.strip()
            
            amt_col = None
            for col in df_source.columns:
                if "待拆分" in col: amt_col = col; break
            if not amt_col: st.error("❌ 费用表中未找到包含‘待拆分金额’的列。"); st.stop()
                
            df_ratio_raw = pd.read_excel(ratio_file, header=None)
            ratio_header_idx = 0
            for idx, row in df_ratio_raw.iterrows():
                if "成本中心编号" in row.values: ratio_header_idx = idx; break
            top_headers = df_ratio_raw.iloc[ratio_header_idx - 1].tolist() if ratio_header_idx > 0 else df_ratio_raw.iloc[0].tolist()
            text_headers = df_ratio_raw.iloc[ratio_header_idx].tolist()
            
            proj_text_to_code = {}
            fixed_names = ['成本中心编号', '成本中心名称', '大部门分类', '项目', '分摊逻辑', '分摊过渡部门', '分摊类型', '合计']
            for t_name, top_code in zip(text_headers, top_headers):
                if pd.notna(t_name) and pd.notna(top_code):
                    t_str = str(t_name).strip(); c_str = str(top_code).strip()
                    if t_str not in fixed_names and not t_str.startswith('Unnamed:'):
                        if c_str.endswith('.0'): c_str = c_str.split('.')[0].zfill(3)
                        else: c_str = c_str.zfill(3)
                        proj_text_to_code[t_str] = c_str
            df_ratio = pd.read_excel(ratio_file, skiprows=ratio_header_idx)
            df_ratio.columns = df_ratio.columns.astype(str).str.strip()
            
            # 🌟🌟【已修改】更新为您的最新描述表述
            st.success(f"✅ 文件加载就绪！当前做账主体为: 【{selected_company}】")
            
            # 🌟🌟【已修改】更新为您的最新按钮文字表述
            if st.button("🚀 开始全自动重分类并导出金蝶Excel"):
                project_cols = list(proj_text_to_code.keys())
                df_source['待拆分金额_numeric'] = pd.to_numeric(df_source[amt_col], errors='coerce')
                df_to_split = df_source[df_source['待拆分金额_numeric'].notna() & (df_source['待拆分金额_numeric'] != 0)]
                
                base_old_info = {
                    'FBillHead(GL_VOUCHER)': 1, 'FAccountBookID': comp_info['book_id'], 'FAccountBookID#Name': selected_company,
                    'FDate': voucher_date, 'FBUSDATE': voucher_date, 'FYEAR': int(current_year), 'FPERIOD': int(current_period),
                    'FVOUCHERGROUPID': 'PRE001', 'FVOUCHERGROUPID#Name': '记', 'FVOUCHERGROUPNO': '1', 'FATTACHMENTS': 0, 'FISADJUSTVOUCHER': '否',
                    'FACCBOOKORGID': comp_info['org_id'], 'FACCBOOKORGID#Name': selected_company,
                    'FCURRENCYID': chosen_currency['id'], 'FCURRENCYID#Name': chosen_currency['name'], 'FEXCHANGERATETYPE': 'HLTX01_SYS', 'FEXCHANGERATETYPE#Name': '固定汇率', 'FEXCHANGERATE': 1
                }
                new_rows = []
                entry_idx = 1
                
                for idx, row in df_to_split.iterrows():
                    try:
                        sub_code = str(row['科目编码']).strip(); sub_name = str(row['科目名称']).strip(); dim_code = str(row['核算维度编码']).strip()
                        if sub_code == '科目编码' or sub_code == 'nan': continue
                        orig_amt = float(row['待拆分金额_numeric'])
                        cc_code = None
                        parts = dim_code.split('/')
                        for p in parts:
                            p_clean = p.strip()
                            if p_clean in df_ratio['成本中心编号'].astype(str).str.strip().values: cc_code = p_clean; break
                        if not cc_code and dim_code in df_ratio['成本中心编号'].astype(str).str.strip().values: cc_code = dim_code
                        if not cc_code: continue
                            
                        ratio_row_data = df_ratio[df_ratio['成本中心编号'].astype(str).str.strip() == cc_code].iloc[0]
                        valid_projects = []
                        for proj_text in project_cols:
                            val = ratio_row_data[proj_text]
                            try:
                                val_float = float(val)
                                if val_float > 0: valid_projects.append({'proj_text': proj_text, 'proj_code': proj_text_to_code[proj_text], 'ratio': val_float})
                            except: pass
                        if not valid_projects: continue
                        df_valid_proj = pd.DataFrame(valid_projects).sort_values(by='ratio', ascending=False)
                        
                        # 借方负数冲销行
                        neg_row = [None] * len(tech_headers)
                        if entry_idx == 1:
                            for k, v in base_old_info.items():
                                if k in tech_headers: neg_row[tech_headers.index(k)] = v
                        neg_row[tech_headers.index('FEntity')] = entry_idx
                        neg_row[tech_headers.index('FEXPLANATION')] = f"重分类-冲原公摊-{sub_name}"
                        neg_row[tech_headers.index('FACCOUNTID')] = sub_code; neg_row[tech_headers.index('FACCOUNTID#Name')] = sub_name
                        neg_row[tech_headers.index('FDEBIT')] = -orig_amt
                        neg_row[tech_headers.index('FCREDIT')] = None; neg_row[tech_headers.index('FAMOUNTFOR')] = None
                        for field in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']: neg_row[tech_headers.index(field)] = base_old_info[field]
                        if 'FDetailID#FF100002' in tech_headers: neg_row[tech_headers.index('FDetailID#FF100002')] = '006'
                        if 'FDetailID#FFlex5' in tech_headers: neg_row[tech_headers.index('FDetailID#FFlex5')] = cc_code
                        new_rows.append(neg_row)
                        entry_idx += 1
                        
                        # 借方正数分配行
                        allocated_sum = 0.0
                        for i, p_row in enumerate(df_valid_proj.itertuples()):
                            is_last = (i == len(df_valid_proj) - 1); current_ratio = p_row.ratio
                            if is_last: amt = round(orig_amt - allocated_sum, 2)
                            else: amt = round(orig_amt * current_ratio, 2); allocated_sum += amt
                            if amt == 0: continue
                            pos_row = [None] * len(tech_headers)
                            pos_row[tech_headers.index('FEntity')] = entry_idx
                            pos_row[tech_headers.index('FEXPLANATION')] = f"重分类-项目分摊-{sub_name}"
                            pos_row[tech_headers.index('FACCOUNTID')] = sub_code; pos_row[tech_headers.index('FACCOUNTID#Name')] = sub_name
                            pos_row[tech_headers.index('FDEBIT')] = amt
                            pos_row[tech_headers.index('FCREDIT')] = None; pos_row[tech_headers.index('FAMOUNTFOR')] = None
                            for field in ['FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name', 'FEXCHANGERATE']: pos_row[tech_headers.index(field)] = base_old_info[field]
                            if 'FDetailID#FF100002' in tech_headers: pos_row[tech_headers.index('FDetailID#FF100002')] = p_row.proj_code
                            if 'FDetailID#FFlex5' in tech_headers: pos_row[tech_headers.index('FDetailID#FFlex5')] = cc_code
                            new_rows.append(pos_row)
                            entry_idx += 1
                    except: pass
                if new_rows:
                    final_df = pd.DataFrame([tech_headers, cn_headers] + new_rows)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer: final_df.to_excel(writer, index=False, header=False, sheet_name='凭证#单据头(FBillHead)')
                    st.download_button(label="📥 下载重分类引入凭证Excel", data=output.getvalue(), file_name=f"金蝶云星空重分类凭证-{selected_company}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e: st.error(f"发生意外错误: {e}")

# ----------------------------------------------------
# 🎛️ 密封舱分流执行网关
# ----------------------------------------------------
if main_mode == "💳 信用卡对账单智能理账(自用)":
    run_credit_card_tool()
else:
    run_reclassification_tool()

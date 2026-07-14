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
    "Crazy Maple Studio Inc": {"book_id": "100", "org_id": "002", "currency_id": "PRE007", "currency_name": "美元"},
    "CRAZY MAPLE  SERVICE  COMPANY": {"book_id": "101", "org_id": "003", "currency_id": "PRE007", "currency_name": "美元"},
    "Crazy Maple  Canada  Inc": {"book_id": "102", "org_id": "004", "currency_id": "PRE008", "currency_name": "加币"},
    "Maple House Inc": {"book_id": "103", "org_id": "005", "currency_id": "PRE007", "currency_name": "美元"},
    "CRAZY MAPLE  INTERACTIVE  HOLDING LTD": {"book_id": "104", "org_id": "006", "currency_id": "PRE007", "currency_name": "美元"},
    "SPICY MAPLE  LIMITED": {"book_id": "105", "org_id": "007", "currency_id": "PRE007", "currency_name": "美元"},
    "CRAZY MAPLE  STUDIO  HK LIMITED": {"book_id": "106", "org_id": "008", "currency_id": "PRE002", "currency_name": "香港元"},
    "New Leaf  Publishing  Inc": {"book_id": "107", "org_id": "009", "currency_id": "PRE007", "currency_name": "美元"},
    "北京枫悦互动科技有限公司": {"book_id": "108", "org_id": "010", "currency_id": "PRE001", "currency_name": "人民币"},
    "深圳枫叶互动科技有限公司": {"book_id": "109", "org_id": "011", "currency_id": "PRE001", "currency_name": "人民币"},
    "杭州枫叶互动科技有限公司": {"book_id": "110", "org_id": "012", "currency_id": "PRE001", "currency_name": "人民币"},
    "B25 LIMITED": {"book_id": "111", "org_id": "013", "currency_id": "PRE007", "currency_name": "美元"},
    "海南枫悦互动科技有限公司": {"book_id": "112", "org_id": "014", "currency_id": "PRE001", "currency_name": "人民币"},
    "北京枫悦互动科技有限公司工会委员会": {"book_id": "113", "org_id": "015", "currency_id": "PRE001", "currency_name": "人民币"},
    "深圳市星尘游戏科技有限公司": {"book_id": "114", "org_id": "016", "currency_id": "PRE001", "currency_name": "人民币"},
    "ReelShort Japan Co., Ltd.": {"book_id": "115", "org_id": "017", "currency_id": "PRE004", "currency_name": "日本日圆"},
    "SWEET MAPLE LIMITED": {"book_id": "116", "org_id": "018", "currency_id": "PRE007", "currency_name": "美元"},
    "深圳枫悦互动科技有限公司": {"book_id": "117", "org_id": "019", "currency_id": "PRE001", "currency_name": "人民币"}
}

# 支持的标准币种字典清单（完全基于官方档案）
CURRENCY_OPTIONS = {
    "人民币 (PRE001)": {"id": "PRE001", "name": "人民币"},
    "美元 (PRE007)": {"id": "PRE007", "name": "美元"},
    "香港元 (PRE002)": {"id": "PRE002", "name": "香港元"},
    "加币 (PRE008)": {"id": "PRE008", "name": "加币"},
    "日本日圆 (PRE004)": {"id": "PRE004", "name": "日本日圆"}
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

# 2. 动态币别选择框
default_curr_label = f

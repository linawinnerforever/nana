import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io

st.set_page_config(page_title="投放费用数据智能汇总工具", layout="wide")

st.title("📊 投放费用月度数据汇总与透视工具")
st.markdown("请在下方上传本月收到的各业务计提表（支持多选），系统将自动合并清洗，并生成左侧明细（指定字段）、右侧透视的专属 Excel 报表。")

# 配置文件上传器
uploaded_files = st.file_uploader(
    "上传业务计提表 (可多选 Excel 文件)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

# 映射配置：根据文件名识别买量产品和默认核算主体
PRODUCT_MAPPING = {
    'chapters': {'product': 'CHAPTERS', 'entity': 'CM'},
    'kiss': {'product': 'KISS', 'entity': 'MH'},
    'maxdrama': {'product': 'MaxDrama', 'entity': 'NL'}, # 用户要求：MaxDrama核算主体改为NL
    'merge': {'product': 'Merge', 'entity': 'CM'},
    'reelshort': {'product': 'Reelshort', 'entity': 'NL'},
    'rsnovel': {'product': 'RS N', 'entity': 'NL'}
}

def process_data(files):
    all_data = []
    
    for f in files:
        fname = f.name.lower()
        prod_info = None
        for key, val in PRODUCT_MAPPING.items():
            if key in fname:
                prod_info = val
                break
        
        if not prod_info:
            st.warning(f"无法识别文件 {f.name} 的产品类别，将跳过此文件。")
            continue
            
        try:
            # 智能探测表头位置
            df_check = pd.read_excel(f, nrows=10)
            header_row = 0
            for i, row in df_check.iterrows():
                if '投放渠道' in row.values or 'spent' in row.values:
                    header_row = i + 1
                    break
            
            df = pd.read_excel(f, skiprows=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            
            if 'spent' not in df.columns:
                st.error(f"文件 {f.name} 中未找到 'spent' 列，请检查格式。")
                continue
                
            # 过滤无效行
            df = df[df['spent'].notna()]
            df = df[pd.to_numeric(df['spent'], errors='coerce').notna()]
            df = df[df['spent'] > 0]
            if '投放渠道' in df.columns:
                df = df[~df['投放渠道'].astype(str).str.contains('合计|Total', na=False)]
            
            # 构建用户要求的左侧表精简字段：投放渠道、开户方、广告户名、spent、买量产品、主体
            processed_df = pd.DataFrame()
            processed_df['投放渠道'] = df['投放渠道'] if '投放渠道' in df.columns else ''
            processed_df['开户方'] = df['开户方'] if '开户方' in df.columns else (df['开户服务商'] if '开户服务商' in df.columns else '')
            processed_df['广告户名'] = df['广告户名'] if '广告户名' in df.columns else ''
            processed_df['spent'] = df['spent'].astype(float)
            processed_df['买量产品'] = prod_info['product']
            processed_df['核算主体'] = prod_info['entity']
            
            all_data.append(processed_df)
        except Exception as e:
            st.error(f"处理文件 {f.name} 时出现错误: {str(e)}")
            
    if not all_data:
        return None
        
    return pd.concat(all_data, ignore_index=True)

if uploaded_files:
    with st.spinner("正在深度清洗并合并数据中..."):
        df_detail = process_data(uploaded_files)
        
    if df_detail is not None:
        st.success("🎉 数据合并清洗成功！")
        
        # 前端预览
        st.subheader("📋 汇总明细数据预览 (左侧结构)")
        st.dataframe(df_detail, use_container_width=True)
        
        # 核心逻辑：右侧透视表字段（买量产品、核算主体、spent、投放渠道、开户方）
        # 先按照指定的聚合维度分组，并将 spent 进行求和
        df_pivot = df_detail.groupby(['买量产品', '核算主体', '投放渠道', '开户方'], as_index=False)['spent'].sum()
        # 调整透视表的列顺序，使其完全符合用户要求的：买量产品、核算主体、spent、投放渠道、开户方
        df_pivot = df_pivot

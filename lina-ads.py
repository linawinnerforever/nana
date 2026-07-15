import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io

st.set_page_config(page_title="投放费用数据智能汇总工具", layout="wide")

st.title("📊 投放费用月度数据汇总与透视工具")
st.markdown("请在下方上传本月收到的各业务计提表（支持多选），系统将自动合并清洗，并生成左侧明细、右侧透视的专属 Excel 报表。")

# 1. 配置文件上传器
uploaded_files = st.file_uploader(
    "上传业务计提表 (可多选 Excel 文件)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

# 映射配置：根据文件名识别买量产品和默认核算主体
PRODUCT_MAPPING = {
    'chapters': {'product': 'CHAPTERS', 'entity': 'CM'},
    'kiss': {'product': 'KISS', 'entity': 'MH'},
    'maxdrama': {'product': 'MaxDrama', 'entity': 'NL'}, # 用户特别要求：MaxDrama改为NL
    'merge': {'product': 'Merge', 'entity': 'CM'},
    'reelshort': {'product': 'Reelshort', 'entity': 'NL'},
    'rsnovel': {'product': 'RS N', 'entity': 'NL'}
}

def process_data(files):
    all_data = []
    
    for f in files:
        fname = f.name.lower()
        # 识别当前文件属于哪个产品
        prod_info = None
        for key, val in PRODUCT_MAPPING.items():
            if key in fname:
                prod_info = val
                break
        
        if not prod_info:
            st.warning(f"无法识别文件 {f.name} 的产品类别，将跳过此文件。")
            continue
            
        # 读取 Excel，通常从包含“投放渠道”的行开始，这里用 skiprows 智能兼容
        try:
            # 先读取前几行探测表头位置
            df_check = pd.read_excel(f, nrows=10)
            header_row = 0
            for i, row in df_check.iterrows():
                if '投放渠道' in row.values or 'spent' in row.values:
                    header_row = i + 1
                    break
            
            df = pd.read_excel(f, skiprows=header_row)
            # 清理列名空格
            df.columns = [str(c).strip() for c in df.columns]
            
            # 如果没找到 spent 列则跳过
            if 'spent' not in df.columns:
                st.error(f"文件 {f.name} 中未找到 'spent' 列，请检查格式。")
                continue
                
            # 过滤掉 spent 为空或非数字的行，以及合计行
            df = df[df['spent'].notna()]
            df = df[pd.to_numeric(df['spent'], errors='coerce').notna()]
            df = df[df['spent'] > 0]
            if '投放渠道' in df.columns:
                df = df[~df['投放渠道'].astype(str).str.contains('合计|Total', na=False)]
            
            # 提取与构建需要的列
            processed_df = pd.DataFrame()
            processed_df['期间'] = [f"{pd.Timestamp.now().month}月-{str(pd.Timestamp.now().year)[2:]}"] * len(df)
            processed_df['投放渠道'] = df['投放渠道'] if '投放渠道' in df.columns else ''
            processed_df['开户服务商'] = df['开户方'] if '开户方' in df.columns else (df['开户服务商'] if '开户服务商' in df.columns else '')
            processed_df['广告户名'] = df['广告户名'] if '广告户名' in df.columns else ''
            
            # 类别与代投服务商逻辑
            if '代投方' in df.columns:
                processed_df['类别'] = df['代投方'].apply(lambda x: '自投' if '自投' in str(x) else '代投')
                processed_df['代投服务商'] = df['代投方'].apply(lambda x: '' if '自投' in str(x) else x)
            else:
                processed_df['类别'] = '自投'
                processed_df['代投服务商'] = ''
                
            processed_df['消耗'] = df['spent'].astype(float)
            processed_df['代投费'] = df['代投费'].astype(float) if '代投费' in df.columns else 0.0
            processed_df['代投费'] = processed_df['代投费'].fillna(0.0)
            
            # 核心业务规则：投放待结算 = spent（消耗）
            processed_df['投放待结算'] = processed_df['消耗']
            
            # 填入买量产品和核算主体
            processed_df['买量产品'] = prod_info['product']
            processed_df['核算主体'] = prod_info['entity']
            
            all_data.append(processed_df)
        except Exception as e:
            st.error(f"处理文件 {f.name} 时出现错误: {str(e)}")
            
    if not all_data:
        return None
        
    final_df = pd.concat(all_data, ignore_index=True)
    return final_df

if uploaded_files:
    with st.spinner("正在深度清洗并合并数据中..."):
        df_detail = process_data(uploaded_files)
        
    if df_detail is not None:
        st.success("🎉 数据合并清洗成功！")
        
        # 在前端展示预览明细
        st.subheader("📋 汇总明细数据预览")
        st.dataframe(df_detail, use_container_width=True)
        
        # 自动生成透视数据（按买量产品、核算主体、投放渠道透视消耗和待结算金额）
        df_pivot = df_detail.groupby(['买量产品', '核算主体', '投放渠道'], as_index=False)[['消耗', '投放待结算']].sum()
        
        # 2. 使用 openpyxl 进行炫酷的左右并排排版排版
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "费用汇总及透视表"
        ws.views.sheetView[0].showGridLines = True # 确保网格线可见
        
        # 样式定义
        font_title = Font(name="微软雅黑", size=14, bold=True, color="FFFFFF")
        font_header = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        font_body = Font(name="微软雅黑", size=10)
        font_total = Font(name="微软雅黑", size=10, bold=True)
        
        fill_detail_title = PatternFill(start_color="2B4C7E", end_color="2B4C7E", fill_type="solid") # 深蓝
        fill_pivot_title = PatternFill(start_color="1E6B52", end_color="1E6B52", fill_type="solid")  # 深绿
        fill_detail_hdr = PatternFill(start_color="4A7BB0", end_color="4A7BB0", fill_type="solid")   # 浅蓝
        fill_pivot_hdr = PatternFill(start_color="339977", end_color="339977", fill_type="solid")    # 浅绿
        fill_zebra = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")        # 斑马纹
        
        thin_border = Border(
            left=Side(style='thin', color='DDDDDD'),
            right=Side(style='thin', color='DDDDDD'),
            top=Side(style='thin', color='DDDDDD'),
            bottom=Side(style='thin', color='DDDDDD')
        )
        double_bottom_border = Border(
            top=Side(style='thin', color='000000'),
            bottom=Side(style='double', color='000000')
        )
        
        # A. 写入左侧明细表表头大标题 (合并 A1:K1)
        ws.merge_cells("A1:K1")
        ws["A1"] = "投放费用清洗明细表"
        ws["A1"].font = font_title
        ws["A1"].fill = fill_detail_title
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        
        # B. 写入右侧透视表表头大标题 (合并 M1:Q1)
        ws.merge_cells("M1:Q1")
        ws["M1"] = "投放费用智能透视表"
        ws["M1"].font = font_title
        ws["M1"].fill = fill_pivot_title
        ws["M1"].alignment = Alignment(horizontal="center", vertical="center")
        
        ws.row_dimensions[1].height = 35
        ws.row_dimensions[2].height = 24
        
        # C. 写入左侧明细列名 (第2行)
        detail_cols = list(df_detail.columns)
        for col_idx, col_name in enumerate(detail_cols, start=1):
            cell = ws.cell(row=2, column=col_idx, value=col_name)
            cell.font = font_header
            cell.fill = fill_detail_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # D. 写入右侧透视列名 (第2行)
        pivot_cols = list(df_pivot.columns)
        for col_idx, col_name in enumerate(pivot_cols, start=13): # 从第13列(M列)开始
            cell = ws.cell(row=2, column=col_idx, value=col_name)
            cell.font = font_header
            cell.fill = fill_pivot_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # E. 填充左侧明细数据
        for row_idx, r in enumerate(dataframe_to_rows(df_detail, index=False, header=False), start=3):
            ws.row_dimensions[row_idx].height = 20
            for col_idx, val in enumerate(r, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                # 针对金额列进行千分位格式化
                if detail_cols[col_idx-1] in ['消耗', '代投费', '投放待结算']:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                if row_idx % 2 == 0:
                    cell.fill = fill_zebra
                    
        # F. 填充右侧透视数据
        for row_idx, r in enumerate(dataframe_to_rows(df_pivot, index=False, header=False), start=3):
            ws.row_dimensions[row_idx].height = 20
            for col_idx, val in enumerate(r, start=13):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                if pivot_cols[col_idx-13] in ['消耗', '投放待结算']:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    
        # G. 添加明细表合计行
        detail_last_row = len(df_detail) + 3
        ws.cell(row=detail_last_row, column=1, value="合计").font = font_total
        ws.cell(row=detail_last_row, column=1).alignment = Alignment(horizontal="center")
        for c_idx in range(2, 12):
            c_cell = ws.cell(row=detail_last_row, column=c_idx)
            c_cell.border = double_bottom_border
            if detail_cols[c_idx-1] in ['消耗', '代投费', '投放待结算']:
                col_letter = openpyxl.utils.get_column_letter(c_idx)
                c_cell.value = f"=SUM({col_letter}3:{col_letter}{detail_last_row-1})"
                c_cell.font = font_total
                c_cell.number_format = '#,##0.00'
                c_cell.alignment = Alignment(horizontal="right")
                
        # H. 添加透视表合计行
        pivot_last_row = len(df_pivot) + 3
        ws.cell(row=pivot_last_row, column=13, value="总计").font = font_total
        ws.cell(row=pivot_last_row, column=13).alignment = Alignment(horizontal="center")
        for c_idx in range(14, 18):
            c_cell = ws.cell(row=pivot_last_row, column=c_idx)
            c_cell.border = double_bottom_border
            if pivot_cols[c_idx-13] in ['消耗', '投放待结算']:
                col_letter = openpyxl.utils.get_column_letter(c_idx)
                c_cell.value = f"=SUM({col_letter}3:{col_letter}{pivot_last_row-1})"
                c_cell.font = font_total
                c_cell.number_format = '#,##0.00'
                c_cell.alignment = Alignment(horizontal="right")

        # 智能微调列宽
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        ws.column_dimensions['L'].width = 5 # 给左右两张表留出隔离空隙
        
        # 导出为二进制流供 Streamlit 下载
        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)
        
        st.markdown("---")
        st.subheader("🚀 导出生成报表")
        st.download_button(
            label="点击下载「左明细右透视」精美Excel报表",
            data=excel_data,
            file_name="🤝投放费用汇总明细与透视表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

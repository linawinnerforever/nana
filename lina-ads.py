import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io

st.set_page_config(page_title="投放费用数据智能汇总工具", layout="wide")

st.title("📊 投放费用月度数据汇总与透视工具 V3")
st.markdown("已更新：总计行已移至数据上方第一行，并采用 `SUBTOTAL` 公式，确保明细与透视金额绝对一致，且支持筛选联动。")

# 配置文件上传器
uploaded_files = st.file_uploader(
    "上传业务计提表 (可多选 Excel 文件)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

# 映射配置
PRODUCT_MAPPING = {
    'chapters': {'product': 'CHAPTERS', 'entity': 'CM'},
    'kiss': {'product': 'KISS', 'entity': 'MH'},
    'maxdrama': {'product': 'MaxDrama', 'entity': 'NL'}, # MaxDrama核算主体改为NL
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
            continue
            
        try:
            df_check = pd.read_excel(f, nrows=10)
            header_row = 0
            for i, row in df_check.iterrows():
                if '投放渠道' in row.values or 'spent' in row.values:
                    header_row = i + 1
                    break
            
            df = pd.read_excel(f, skiprows=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            
            if 'spent' not in df.columns:
                continue
                
            # 数据清洗与过滤
            df = df[df['spent'].notna()]
            df = df[pd.to_numeric(df['spent'], errors='coerce').notna()]
            df = df[df['spent'] > 0]
            if '投放渠道' in df.columns:
                df = df[~df['投放渠道'].astype(str).str.contains('合计|Total', na=False)]
            
            processed_df = pd.DataFrame()
            processed_df['投放渠道'] = df['投放渠道'] if '投放渠道' in df.columns else ''
            processed_df['开户方'] = df['开户方'] if '开户方' in df.columns else (df['开户服务商'] if '开户服务商' in df.columns else '')
            processed_df['广告户名'] = df['广告户名'] if '广告户名' in df.columns else ''
            processed_df['spent'] = df['spent'].astype(float)
            processed_df['买量产品'] = prod_info['product']
            processed_df['核算主体'] = prod_info['entity']
            
            all_data.append(processed_df)
        except Exception as e:
            st.error(f"处理文件 {f.name} 时出错: {str(e)}")
            
    return pd.concat(all_data, ignore_index=True) if all_data else None

if uploaded_files:
    with st.spinner("正在处理数据中..."):
        df_detail = process_data(uploaded_files)
        
    if df_detail is not None:
        st.success("🎉 数据合并清洗成功！")
        
        # 构造右侧透视数据
        df_pivot = df_detail.groupby(['买量产品', '核算主体', '投放渠道', '开户方'], as_index=False)['spent'].sum()
        df_pivot = df_pivot[['买量产品', '核算主体', 'spent', '投放渠道', '开户方']]
        
        # 使用 openpyxl 进行排版
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "费用汇总及透视表"
        ws.views.sheetView[0].showGridLines = True
        
        # 样式定义
        font_title = Font(name="微软雅黑", size=12, bold=True, color="FFFFFF")
        font_header = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        font_total = Font(name="微软雅黑", size=10, bold=True, color="FF5722") # 橙色高亮总计
        font_body = Font(name="微软雅黑", size=10)
        
        fill_detail_title = PatternFill(start_color="2B4C7E", end_color="2B4C7E", fill_type="solid")
        fill_pivot_title = PatternFill(start_color="1E6B52", end_color="1E6B52", fill_type="solid")
        fill_detail_hdr = PatternFill(start_color="4A7BB0", end_color="4A7BB0", fill_type="solid")
        fill_pivot_hdr = PatternFill(start_color="339977", end_color="339977", fill_type="solid")
        fill_total = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid") # 总计行浅橙色背景
        fill_zebra = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'),
            top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD')
        )
        total_border = Border(
            top=Side(style='thin', color='FF5722'), bottom=Side(style='double', color='FF5722'),
            left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD')
        )
        
        # 1. 写入第1行：大标题
        ws.merge_cells("A1:F1")
        ws["A1"] = "投放费用明细表"
        ws["A1"].font = font_title
        ws["A1"].fill = fill_detail_title
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells("H1:L1")
        ws["H1"] = "投放费用透视表"
        ws["H1"].font = font_title
        ws["H1"].fill = fill_pivot_title
        ws["H1"].alignment = Alignment(horizontal="center", vertical="center")
        
        # 2. 写入第2行：列名表头
        detail_cols = list(df_detail.columns)
        for col_idx, col_name in enumerate(detail_cols, start=1):
            cell = ws.cell(row=2, column=col_idx, value=col_name)
            cell.font = font_header
            cell.fill = fill_detail_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        pivot_cols = list(df_pivot.columns)
        for col_idx, col_name in enumerate(pivot_cols, start=8):
            cell = ws.cell(row=2, column=col_idx, value=col_name)
            cell.font = font_header
            cell.fill = fill_pivot_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # 3. 写入第3行：【新增】数据列上方的 SUBTOTAL 总计行
        detail_end_row = len(df_detail) + 3
        pivot_end_row = len(df_pivot) + 3
        
        # 左侧明细表 SUBTOTAL (D列为spent)
        ws.cell(row=3, column=1, value="总计 (SUBTOTAL)").font = font_total
        ws.cell(row=3, column=1).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 7):
            cell = ws.cell(row=3, column=c)
            cell.fill = fill_total
            cell.border = total_border
            if detail_cols[c-1] == 'spent':
                cell.value = f"=SUBTOTAL(9, D4:D{detail_end_row})"
                cell.font = font_total
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
                
        # 右侧透视表 SUBTOTAL (J列为spent)
        ws.cell(row=3, column=8, value="总计 (SUBTOTAL)").font = font_total
        ws.cell(row=3, column=8).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(8, 13):
            cell = ws.cell(row=3, column=c)
            cell.fill = fill_total
            cell.border = total_border
            if pivot_cols[c-8] == 'spent':
                cell.value = f"=SUBTOTAL(9, J4:J{pivot_end_row})"
                cell.font = font_total
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
        
        # 4. 写入明细数据 (从第4行开始)
        for row_idx, r in enumerate(dataframe_to_rows(df_detail, index=False, header=False), start=4):
            for col_idx, val in enumerate(r, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                if detail_cols[col_idx-1] == 'spent':
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                if row_idx % 2 == 0:
                    cell.fill = fill_zebra
                    
        # 5. 写入透视数据 (从第4行开始)
        for row_idx, r in enumerate(dataframe_to_rows(df_pivot, index=False, header=False), start=4):
            for col_idx, val in enumerate(r, start=8):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                if pivot_cols[col_idx-8] == 'spent':
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # 行高与列宽调整
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 22
        ws.row_dimensions[3].height = 24 # 总计行略高
        
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        ws.column_dimensions['G'].width = 5 # 隔离带
        
        # 导出数据流
        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)
        
        st.markdown("---")
        st.subheader("🚀 导出新版报表")
        st.download_button(
            label="点击下载「顶部SUBTOTAL总计版」Excel报表",
            data=excel_data,
            file_name="🤝投放费用汇总_顶部SUBTOTAL版.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

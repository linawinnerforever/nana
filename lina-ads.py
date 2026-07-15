import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io

st.set_page_config(page_title="投放费用数据智能汇总工具", layout="wide")

st.title("📊 投放费用月度数据汇总与透视工具 V8")
st.markdown("说明：保留了原有的报表下载；同时根据您的底表，额外独立生成了一张专供领导的高管级汇总 Excel 表。")

uploaded_files = st.file_uploader("上传业务计提表 (可多选 Excel 文件)", type=["xlsx", "xls"], accept_multiple_files=True)

PRODUCT_MAPPING = {
    'chapters': {'product': 'CHAPTERS', 'entity': 'CM'},
    'kiss': {'product': 'KISS', 'entity': 'MH'},
    'maxdrama': {'product': 'MaxDrama', 'entity': 'NL'},
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
                
            df = df[df['spent'].notna()]
            df['spent_num'] = pd.to_numeric(df['spent'], errors='coerce')
            df = df[df['spent_num'].notna()]
            df = df[df['spent_num'] > 0]
            if '投放渠道' in df.columns:
                df = df[~df['投放渠道'].astype(str).str.contains('合计|Total', na=False)]
            
            processed_df = pd.DataFrame()
            processed_df['投放渠道'] = df['投放渠道'].fillna('').astype(str).str.strip()
            
            if '开户方' in df.columns:
                processed_df['开户方'] = df['开户方'].fillna('').astype(str).str.strip()
            elif '开户服务商' in df.columns:
                processed_df['开户方'] = df['开户服务商'].fillna('').astype(str).str.strip()
            else:
                processed_df['开户方'] = ''
                
            processed_df['广告户名'] = df['广告户名'].fillna('').astype(str).str.strip() if '广告户名' in df.columns else ''
            processed_df['spent'] = df['spent_num']
            processed_df['买量产品'] = prod_info['product']
            processed_df['核算主体'] = prod_info['entity']
            
            all_data.append(processed_df)
        except Exception as e:
            st.error(f"处理文件 {f.name} 时出错: {str(e)}")
            
    return pd.concat(all_data, ignore_index=True) if all_data else None

if uploaded_files:
    with st.spinner("正在汇总数据..."):
        df_detail = process_data(uploaded_files)
        
    if df_detail is not None:
        st.success("🎉 数据合并完成！请在下方选择您需要下载的 Excel 文件：")
        
        # 基础透视数据准备
        df_pivot = df_detail.groupby(['买量产品', '核算主体', '投放渠道', '开户方'], as_index=False, dropna=False)['spent'].sum()
        df_pivot = df_pivot[['买量产品', '核算主体', 'spent', '投放渠道', '开户方']]
        
        # ==========================================
        # 按钮一：原有 Excel 报表生成逻辑（不做任何修改）
        # ==========================================
        wb_orig = openpyxl.Workbook()
        ws_orig = wb_orig.active
        ws_orig.title = "费用汇总及透视表"
        ws_orig.views.sheetView[0].showGridLines = True
        
        font_title = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        font_header = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        font_total = Font(name="微软雅黑", size=11, bold=True, color="D32F2F")
        font_body = Font(name="微软雅黑", size=10)
        
        fill_detail_title = PatternFill(start_color="2B4C7E", end_color="2B4C7E", fill_type="solid")
        fill_pivot_title = PatternFill(start_color="1E6B52", end_color="1E6B52", fill_type="solid")
        fill_detail_hdr = PatternFill(start_color="4A7BB0", end_color="4A7BB0", fill_type="solid")
        fill_pivot_hdr = PatternFill(start_color="339977", end_color="339977", fill_type="solid")
        fill_zebra = PatternFill(start_color="F9FBFC", end_color="F9FBFC", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'),
            top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0')
        )
        total_border = Border(
            top=Side(style='medium', color='D32F2F'), bottom=Side(style='medium', color='D32F2F'),
            left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0')
        )
        
        detail_cols = ['投放渠道', '开户方', '广告户名', 'spent', '买量产品', '核算主体']
        pivot_cols = ['买量产品', '核算主体', 'spent', '投放渠道', '开户方']
        
        detail_end = len(df_detail) + 3
        pivot_end = len(df_pivot) + 3
        
        ws_orig.cell(row=1, column=1, value="总计 (SUBTOTAL)").font = font_total
        for c in range(1, 7):
            cell = ws_orig.cell(row=1, column=c)
            cell.border = total_border
            if detail_cols[c-1] == 'spent':
                cell.value = f"=SUBTOTAL(9, D4:D{detail_end})"
                cell.font = font_total
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
        ws_orig.cell(row=1, column=8, value="总计 (SUBTOTAL)").font = font_total
        for c in range(8, 13):
            cell = ws_orig.cell(row=1, column=c)
            cell.border = total_border
            if pivot_cols[c-8] == 'spent':
                cell.value = f"=SUBTOTAL(9, J4:J{pivot_end})"
                cell.font = font_total
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c == 8:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
        
        ws_orig.merge_cells("A2:F2")
        ws_orig["A2"] = "投放费用明细表"
        ws_orig["A2"].font = font_title
        ws_orig["A2"].fill = fill_detail_title
        ws_orig["A2"].alignment = Alignment(horizontal="center", vertical="center")
        
        ws_orig.merge_cells("H2:L2")
        ws_orig["H2"] = "投放费用透视表"
        ws_orig["H2"].font = font_title
        ws_orig["H2"].fill = fill_pivot_title
        ws_orig["H2"].alignment = Alignment(horizontal="center", vertical="center")
        
        for idx, col in enumerate(detail_cols, 1):
            cell = ws_orig.cell(row=3, column=idx, value=col)
            cell.font = font_header
            cell.fill = fill_detail_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        for idx, col in enumerate(pivot_cols, 8):
            cell = ws_orig.cell(row=3, column=idx, value=col)
            cell.font = font_header
            cell.fill = fill_pivot_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        for r_idx, row in enumerate(dataframe_to_rows(df_detail[detail_cols], index=False, header=False), start=4):
            for c_idx, val in enumerate(row, start=1):
                cell = ws_orig.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                if detail_cols[c_idx-1] == 'spent':
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                if r_idx % 2 == 0:
                    cell.fill = fill_zebra
                    
        for r_idx, row in enumerate(dataframe_to_rows(df_pivot[pivot_cols], index=False, header=False), start=4):
            for c_idx, val in enumerate(row, start=8):
                cell = ws_orig.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                if pivot_cols[c_idx-8] == 'spent':
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        ws_orig.row_dimensions[1].height = 26
        ws_orig.row_dimensions[2].height = 24
        ws_orig.row_dimensions[3].height = 22
        
        for col in ws_orig.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws_orig.column_dimensions[col_letter].width = max(max_len + 3, 12)
        ws_orig.column_dimensions['G'].width = 4
        
        excel_data_orig = io.BytesIO()
        wb_orig.save(excel_data_orig)
        excel_data_orig.seek(0)

        # ==========================================
        # 按钮二：独立生成“给领导的汇总表” (全新生成独立Excel)
        # ==========================================
        wb_leader = openpyxl.Workbook()
        ws_leader = wb_leader.active
        ws_leader.title = "领导审阅汇总"
        ws_leader.views.sheetView[0].showGridLines = True
        
        # 领导偏好的沉稳大气风格
        font_l_title = Font(name="微软雅黑", size=14, bold=True, color="333333")
        font_l_hdr = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        font_l_body = Font(name="微软雅黑", size=11, color="000000")
        font_l_total = Font(name="微软雅黑", size=11, bold=True, color="B00020")
        
        fill_l_hdr = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid") # 绅士灰蓝
        fill_l_zebra = PatternFill(start_color="F8F9F9", end_color="F8F9F9", fill_type="solid")
        fill_l_total = PatternFill(start_color="EAEDED", end_color="EAEDED", fill_type="solid")
        
        b_thin = Side(style='thin', color='BDC3C7')
        b_double = Side(style='double', color='333333')
        b_thick = Side(style='medium', color='333333')
        
        border_l_cell = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)
        border_l_total = Border(left=b_thin, right=b_thin, top=b_thick, bottom=b_double)

        # 1. 铺设大标题
        ws_leader.cell(row=2, column=2, value="📊 推广费用各主体及产品消耗情况汇总表").font = font_l_title
        
        # 2. 核心聚合数据：给领导看的数据颗粒度通常到 产品 + 核算主体 + 渠道
        df_l_summary = df_detail.groupby(['买量产品', '核算主体', '投放渠道'], as_index=False)['spent'].sum()
        df_l_summary = df_l_summary.sort_values(by=['买量产品', 'spent'], ascending=[True, False])
        
        # ！！在此处定义输出给领导汇总表的列名结构 (可根据您的喜好自由微调文字)！！
        headers_leader = ["买量产品", "核算主体", "投放渠道", "消耗金额 (USD)"]
        
        for c_idx, h_text in enumerate(headers_leader, start=2):
            cell = ws_leader.cell(row=4, column=c_idx, value=h_text)
            cell.font = font_l_hdr
            cell.fill = fill_l_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_l_cell
        ws_leader.row_dimensions[4].height = 28
        
        # 3. 循环写入数据并追加小计
        l_curr_row = 5
        for prod, group in df_l_summary.groupby('买量产品'):
            group_start = l_curr_row
            
            for _, row in group.iterrows():
                ws_leader.cell(row=l_curr_row, column=2, value=row['买量产品']).alignment = Alignment(horizontal="center", vertical="center")
                ws_leader.cell(row=l_curr_row, column=3, value=row['核算主体']).alignment = Alignment(horizontal="center", vertical="center")
                ws_leader.cell(row=l_curr_row, column=4, value=row['投放渠道']).alignment = Alignment(horizontal="center", vertical="center")
                
                cell_amt = ws_leader.cell(row=l_curr_row, column=5, value=row['spent'])
                cell_amt.number_format = '#,##0.00'
                cell_amt.alignment = Alignment(horizontal="right", vertical="center")
                
                for c in range(2, 6):
                    t_cell = ws_leader.cell(row=l_curr_row, column=c)
                    t_cell.font = font_l_body
                    t_cell.border = border_l_cell
                    if l_curr_row % 2 == 0:
                        t_cell.fill = fill_l_zebra
                l_curr_row += 1
                
            # 产品级小计行
            ws_leader.cell(row=l_curr_row, column=2, value=f"{prod} 合计").alignment = Alignment(horizontal="left", vertical="center")
            cell_sub = ws_leader.cell(row=l_curr_row, column=5, value=f"=SUM(E{group_start}:E{l_curr_row-1})")
            cell_sub.number_format = '#,##0.00'
            cell_sub.alignment = Alignment(horizontal="right", vertical="center")
            
            for c in range(2, 6):
                t_cell = ws_leader.cell(row=l_curr_row, column=c)
                t_cell.font = font_l_total
                t_cell.fill = fill_l_total
                t_cell.border = border_l_total
            ws_leader.row_dimensions[l_curr_row].height = 24
            l_curr_row += 2 # 留出空行，使排版更具呼吸感
            
        # 4. 全盘总计行
        ws_leader.cell(row=l_curr_row, column=2, value="总计 (Grand Total)").alignment = Alignment(horizontal="left", vertical="center")
        # 稳健求和：明细值加总
        cell_grand = ws_leader.cell(row=l_curr_row, column=5, value=f"=SUM(E5:E{l_curr_row-2})/2") 
        # 更加直观的做法是直接把刚才所有的小计行抓出来求和
        subtotal_formula_parts = [f"E{r}" for r in range(5, l_curr_row-1) if "合计" in str(ws_leader.cell(row=r, column=2).value)]
        if subtotal_formula_parts:
            cell_grand.value = f"=SUM({','.join(subtotal_formula_parts)})"
            
        cell_grand.number_format = '#,##0.00'
        cell_grand.alignment = Alignment(horizontal="right", vertical="center")
        
        for c in range(2, 6):
            t_cell = ws_leader.cell(row=l_curr_row, column=c)
            t_cell.font = font_l_total
            t_cell.fill = fill_l_total
            t_cell.border = border_l_total
        ws_leader.row_dimensions[l_curr_row].height = 26
        
        # 5. 自动算列宽
        for col in ws_leader.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            if col_letter != 'A':
                ws_leader.column_dimensions[col_letter].width = max(max_len + 5, 16)
        ws_leader.column_dimensions['A'].width = 3
        
        excel_data_leader = io.BytesIO()
        wb_leader.save(excel_data_leader)
        excel_data_leader.seek(0)
        
        # ==========================================
        # Streamlit 界面双按钮独立渲染
        # ==========================================
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.info("常规业务报表（包含您原有的样式与完整的明细/透视）")
            st.download_button(
                label="点击下载新样式 Excel 报表",
                data=excel_data_orig,
                file_name="🤝投放费用汇总_对齐美化版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col2:
            st.success("专属管理层报表（干净利落的高管视角汇总表）")
            st.download_button(
                label="👑 点击下载给领导的汇总表",
                data=excel_data_leader,
                file_name="📊2026年5月推广费用-各主体情况汇总-给领导审阅.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

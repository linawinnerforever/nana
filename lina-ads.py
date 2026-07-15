import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re

st.set_page_config(page_title="投放费用数据智能汇总工具", layout="wide")

st.title("📊 投放费用月度数据汇总与透视工具 V10")
st.markdown("特性：修复了列缺失导致的报错。动态提取文件名中的月份，自动更新『期间』，且给领导的汇总表样式完全克隆原厂底表。")

uploaded_files = st.file_uploader("上传业务计提表 (可多选 Excel 文件)", type=["xlsx", "xls"], accept_multiple_files=True)

PRODUCT_MAPPING = {
    'chapters': {'product': 'CHAPTERS', 'entity': 'CM'},
    'kiss': {'product': 'KISS', 'entity': 'MH'},
    'maxdrama': {'product': 'MaxDrama', 'entity': 'CM'},
    'merge': {'product': 'Merge', 'entity': 'CM'},
    'reelshort': {'product': 'Reelshort', 'entity': 'NL'},
    'rsnovel': {'product': 'RS N', 'entity': 'NL'}
}

def process_data(files):
    all_data = []
    detected_period = "未知期间"
    
    # 尝试从上传的文件名中提取“X月-26”或“X月”
    for f in files:
        match = re.search(r'(\d+年\d+月|\d+月)', f.name)
        if match:
            raw_month = match.group(1) # 例如 "2026年6月" 或 "6月"
            if "年" in raw_month:
                # 转换成类似 "6月-26" 的底表常见规范
                year_part = raw_month.split("年")[0][-2:]
                month_part = raw_month.split("年")[1]
                detected_period = f"{month_part}-{year_part}"
            else:
                detected_period = f"{raw_month}-26"
            break

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
                if '投放渠道' in row.values or 'spent' in row.values or '消耗' in row.values:
                    header_row = i + 1
                    break
            
            df = pd.read_excel(f, skiprows=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            
            # 兼容 spent 或 消耗 列
            amt_col = 'spent' if 'spent' in df.columns else ('消耗' if '消耗' in df.columns else None)
            if not amt_col:
                continue
                
            df = df[df[amt_col].notna()]
            df['spent_num'] = pd.to_numeric(df[amt_col], errors='coerce')
            df = df[df['spent_num'].notna()]
            # 兼容负数退款，不剔除 <= 0 的数据
            
            # 过滤合计行
            channel_col = '投放渠道' if '投放渠道' in df.columns else df.columns[0]
            df = df[~df[channel_col].astype(str).str.contains('合计|Total', na=False)]
            
            processed_df = pd.DataFrame()
            
            # 1. 期间 (优先取底表内部的，如果没有则用文件名里提取出来的)
            if '期间' in df.columns and not df['期间'].dropna().empty:
                processed_df['期间'] = df['期间'].fillna(detected_period)
                detected_period = str(df['期间'].dropna().iloc[0])
            else:
                processed_df['期间'] = detected_period
                
            # 2. 投放渠道
            processed_df['投放渠道'] = df['投放渠道'].fillna('').astype(str).str.strip()
            
            # 3. 开户服务商
            if '开户服务商' in df.columns:
                processed_df['开户服务商'] = df['开户服务商'].fillna('').astype(str).str.strip()
            elif '开户方' in df.columns:
                processed_df['开户服务商'] = df['开户方'].fillna('').astype(str).str.strip()
            else:
                processed_df['开户服务商'] = ''
                
            # 4. 广告户名
            processed_df['广告户名'] = df['广告户名'].fillna('').astype(str).str.strip() if '广告户名' in df.columns else ''
            
            # 5. 类别
            processed_df['类别'] = df['类别'].fillna('自投').astype(str).str.strip() if '类别' in df.columns else '自投'
            
            # 6. 代投服务商
            processed_df['代投服务商'] = df['代投服务商'].fillna('').astype(str).str.strip() if '代投服务商' in df.columns else ''
            
            # 7. 消耗
            processed_df['消耗'] = df['spent_num']
            
            # 8. 代投费
            processed_df['代投费'] = pd.to_numeric(df['代投费'], errors='coerce').fillna(0) if '代投费' in df.columns else 0
            
            # 9. 投放待结算
            if '投放待结算' in df.columns:
                processed_df['投放待结算'] = pd.to_numeric(df['投放待结算'], errors='coerce').fillna(processed_df['消耗'] + processed_df['代投费'])
            else:
                processed_df['投放待结算'] = processed_df['消耗'] + processed_df['代投费']
            
            # 10. 买量产品 与 11. 核算主体
            processed_df['买量产品'] = prod_info['product']
            processed_df['核算主体'] = prod_info['entity']
            
            all_data.append(processed_df)
        except Exception as e:
            st.error(f"处理文件 {f.name} 时出错: {str(e)}")
            
    if all_data:
        full_df = pd.concat(all_data, ignore_index=True)
        return full_df, detected_period
    return None, detected_period

if uploaded_files:
    with st.spinner("正在汇总数据..."):
        df_detail, current_month = process_data(uploaded_files)
        
    if df_detail is not None:
        st.success(f"🎉 数据合并成功！已自动动态识别月份为：【{current_month}】")
        
        # 基础透视数据准备（保留旧版兼容）
        df_pivot = df_detail.groupby(['买量产品', '核算主体', '投放渠道', '开户服务商'], as_index=False, dropna=False)['消耗'].sum()
        df_pivot.rename(columns={'开户服务商': '开户方', '消耗': 'spent'}, inplace=True)
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
            
        for r_idx, row in enumerate(dataframe_to_rows(df_detail.rename(columns={'开户服务商': '开户方', '消耗': 'spent'})[detail_cols], index=False, header=False), start=4):
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
        # 按钮二：【给领导的汇总表】克隆原厂底表结构 (完全无损修复)
        # ==========================================
        wb_leader = openpyxl.Workbook()
        ws_leader = wb_leader.active
        ws_leader.title = "各主体情况汇总"
        ws_leader.views.sheetView[0].showGridLines = True
        
        font_leader_hdr = Font(name="微软雅黑", size=10, bold=True)
        font_leader_body = Font(name="微软雅黑", size=10)
        font_leader_top = Font(name="Arial", size=11, bold=True)
        
        align_center = Alignment(horizontal="center", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        
        leader_thin_border = Border(
            left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC')
        )
        
        # 结构化列名
        leader_headers = ["期间", "投放渠道", "开户服务商", "广告户名", "类别", "代投服务商", "消耗", "代投费", "投放待结算", "买量产品", "核算主体"]
        
        # 直接使用清洗完确保存在的列进行导出，防止 KeyError
        df_leader_final = df_detail[leader_headers].copy()
        
        data_end_row = len(df_leader_final) + 3
        
        # Row 1: 首行放置总金额公式汇总
        ws_leader.cell(row=1, column=7, value=f"=SUM(G4:G{data_end_row})").font = font_leader_top
        ws_leader.cell(row=1, column=7).number_format = '#,##0.00'
        ws_leader.cell(row=1, column=7).alignment = align_right
        
        ws_leader.cell(row=1, column=9, value=f"=SUM(I4:I{data_end_row})").font = font_leader_top
        ws_leader.cell(row=1, column=9).number_format = '#,##0.00'
        ws_leader.cell(row=1, column=9).alignment = align_right
        
        # Row 3: 真正的表头行
        for idx, h_name in enumerate(leader_headers, 1):
            cell = ws_leader.cell(row=3, column=idx, value=h_name)
            cell.font = font_leader_hdr
            cell.alignment = align_center
            cell.border = leader_thin_border
        ws_leader.row_dimensions[3].height = 24
        
        # Row 4 起：逐行填充拉平数据
        for r_idx, row in enumerate(dataframe_to_rows(df_leader_final, index=False, header=False), start=4):
            for c_idx, val in enumerate(row, start=1):
                cell = ws_leader.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_leader_body
                cell.border = leader_thin_border
                
                if leader_headers[c_idx-1] in ["消耗", "代投费", "投放待结算"]:
                    cell.number_format = '#,##0.00'
                    cell.alignment = align_right
                else:
                    cell.alignment = align_center
            ws_leader.row_dimensions[r_idx].height = 20
            
        # 自动调整列宽
        for col in ws_leader.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws_leader.column_dimensions[col_letter].width = max(max_len + 3, 13)
            
        excel_data_leader = io.BytesIO()
        wb_leader.save(excel_data_leader)
        excel_data_leader.seek(0)
        
        # 下载板块呈现
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.info("📊 常规业务分析汇总表")
            st.download_button(
                label="点击下载新样式 Excel 报表",
                data=excel_data_orig,
                file_name=f"🤝投放费用汇总_对齐美化版_{current_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col2:
            st.success("👑 专属管理层汇报汇总表 (已按底表样式完全拉平)")
            st.download_button(
                label="点击下载给领导的汇总表",
                data=excel_data_leader,
                file_name=f"推广费用-各主体情况汇总_{current_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re

st.set_page_config(page_title="投放费用数据智能汇总工具", layout="wide")

st.title("📊 投放费用月度数据汇总与透视工具 V23 (完美凭证版)")
st.markdown("特性：**供应商编码已严格锁定：根据核算主体(CM/MH)与匹配出的『供应商-金蝶』，跨表追溯 MP 表 A列/F列的对应编码。**")

# 提供双文件上传器
col_u1, col_u2 = st.columns(2)
with col_u1:
    uploaded_files = st.file_uploader("1. 上传业务计提表 (可多选 Excel)", type=["xlsx", "xls"], accept_multiple_files=True)
with col_u2:
    mp_file = st.file_uploader("2. 上传最新的投放费用 MP 主数据映射表 (单选 Excel)", type=["xlsx", "xls"])

PRODUCT_MAPPING = {
    'chapters': {'product': 'CHAPTERS', 'entity': 'CM'},
    'kiss': {'product': 'KISS', 'entity': 'MH'},
    'maxdrama': {'product': 'MaxDrama', 'entity': 'NL'}, 
    'merge': {'product': 'Merge', 'entity': 'CM'},
    'reelshort': {'product': 'Reelshort', 'entity': 'NL'},
    'rsnovel': {'product': 'RS N', 'entity': 'NL'}
}

REQUIRED_SHEETS = {
    'Advertising-Chapters': {'key': 'chapters', 'product': 'CHAPTERS', 'entity': 'CM'},
    'Advertising-Kiss': {'key': 'kiss', 'product': 'KISS', 'entity': 'MH'},
    'Advertising-MaxDrama': {'key': 'maxdrama', 'product': 'MaxDrama', 'entity': 'NL'},
    'Advertising-Merge': {'key': 'merge', 'product': 'Merge', 'entity': 'CM'},
    'Advertising-Reelshort': {'key': 'reelshort', 'product': 'Reelshort', 'entity': 'NL'},
    'Advertising-RS N': {'key': 'rsnovel', 'product': 'RS N', 'entity': 'NL'}
}

def clean_amount(val):
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace('$', '').replace(',', '')
    val_str = val_str.replace('−', '-').replace('—', '-')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

@st.cache_data
def load_mp_matrix(file):
    """读取 MP 表，强制全文本以防 Excel 吞掉编码前置零"""
    try:
        df_mp = pd.read_excel(file, skiprows=1, dtype=str)
        df_mp.columns = [str(c).strip() for c in df_mp.columns]
        return df_mp
    except Exception as e:
        st.error(f"解析 MP 映射表失败: {str(e)}")
        return None

def process_data(files):
    all_data = []
    detected_period = "未知期间"
    file_month_label = "X月"
    
    for f in files:
        match = re.search(r'(\d+年\d+月|\d+月)', f.name)
        if match:
            raw_month = match.group(1)
            if "年" in raw_month:
                year_part = raw_month.split("年")[0][-2:]
                month_part = raw_month.split("年")[1]
                detected_period = f"{month_part}-{year_part}" 
                file_month_label = month_part
            else:
                detected_period = f"{raw_month}-26"
                file_month_label = raw_month
            break

    for f in files:
        fname = f.name.lower()
        prod_info = None
        target_sheet_name = None
        
        for s_name, info in REQUIRED_SHEETS.items():
            if info['key'] in fname:
                prod_info = info
                target_sheet_name = s_name
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
            
            amt_col = 'spent' if 'spent' in df.columns else ('消耗' if '消耗' in df.columns else None)
            if not amt_col:
                continue
                
            df = df[df[amt_col].notna()]
            df['spent_num'] = df[amt_col].apply(clean_amount)
            
            channel_col = '投放渠道' if '投放渠道' in df.columns else df.columns[0]
            df = df[~df[channel_col].astype(str).str.contains('合计|Total', na=False)]
            
            processed_df = pd.DataFrame()
            processed_df['期间'] = detected_period
            processed_df['投放渠道'] = df['投放渠道'].fillna('').astype(str).str.strip()
            
            if '开户方' in df.columns:
                processed_df['开户服务商'] = df['开户方'].fillna('').astype(str).str.strip()
            elif '开户服务商' in df.columns:
                processed_df['开户服务商'] = df['开户服务商'].fillna('').astype(str).str.strip()
            else:
                processed_df['开户服务商'] = ''
                
            processed_df['广告户名'] = df['广告户名'].fillna('').astype(str).str.strip() if '广告户名' in df.columns else ''
            processed_df['类别'] = df['类别'].fillna('自投').astype(str).str.strip() if '类别' in df.columns else '自投'
            processed_df['代投服务商'] = df['代投服务商'].fillna('').astype(str).str.strip() if '代投服务商' in df.columns else ''
            
            processed_df['消耗'] = df['spent_num']
            processed_df['代投费'] = df['代投费'].apply(clean_amount) if '代投费' in df.columns else 0.0
            processed_df['投放待结算'] = processed_df['消耗'] + processed_df['代投费']
            
            processed_df['买量产品'] = prod_info['product']
            processed_df['核算主体'] = prod_info['entity']
            processed_df['Target_Sheet'] = target_sheet_name
            
            all_data.append(processed_df)
        except Exception as e:
            st.error(f"处理文件 {f.name} 时出错: {str(e)}")
            
    if all_data:
        full_df = pd.concat(all_data, ignore_index=True)
        return full_df, detected_period, file_month_label
    return None, detected_period, file_month_label

if uploaded_files:
    with st.spinner("正在智能清洗流水明细数据..."):
        df_detail, current_month, month_label = process_data(uploaded_files)
        
    if df_detail is not None:
        df_mp_matrix = None
        if mp_file:
            df_mp_matrix = load_mp_matrix(mp_file)
            
        df_pivot = df_detail.groupby(['买量产品', '核算主体', '投放渠道', '开户服务商'], as_index=False, dropna=False)['消耗'].sum()
        df_pivot.rename(columns={'开户服务商': '开户方', '消耗': 'spent'}, inplace=True)
        
        # 供应商-渠道主键
        df_pivot['供应商-渠道'] = df_pivot.apply(
            lambda r: str(r['开户方']).strip() if str(r['开户方']).strip() != "" else str(r['投放渠道']).strip(), axis=1
        )
        
        df_pivot['项目'] = ""
        df_pivot['供应商-金蝶'] = ""
        df_pivot['供应商编码'] = ""
        df_pivot = df_pivot.astype({'项目': 'object', '供应商-金蝶': 'object', '供应商编码': 'object'})
        
        if df_mp_matrix is not None:
            # 1. 建立项目三位财务编码字典 (N列 -> O列)
            dict_project = {}
            for _, r in df_mp_matrix.iterrows():
                k_prod = str(r.iloc[13]).strip().lower() 
                v_code = str(r.iloc[14]).strip()        
                if k_prod and k_prod != 'nan':
                    dict_project[k_prod] = v_code.split('.')[0].zfill(3)
            
            # 2. 建立『渠道』到『金蝶名称』的桥梁字典 (K列 -> L列)
            dict_supplier_jindie = {}
            for _, r in df_mp_matrix.iterrows():
                k_chan = str(r.iloc[10]).strip().lower() 
                v_jindie = str(r.iloc[11]).strip()       
                if k_chan and k_chan != 'nan':
                    dict_supplier_jindie[k_chan] = v_jindie
            
            # 3. 核心修复逻辑：建立『金蝶名称』到『供应商编码』的安全反查词典
            dict_cm_code_by_name = {}  # CM视角：C列名称 -> A列编码
            dict_mh_code_by_name = {}  # MH视角：H列名称 -> F列编码
            
            for _, r in df_mp_matrix.iterrows():
                cm_name = str(r.iloc[2]).strip()   # C列: CM名称
                cm_code = str(r.iloc[0]).strip()   # A列: CM编码
                mh_name = str(r.iloc[7]).strip()   # H列: MH名称
                mh_code = str(r.iloc[5]).strip()   # F列: MH编码
                
                if cm_name and cm_name != 'nan':
                    dict_cm_code_by_name[cm_name.lower()] = cm_code
                if mh_name and mh_name != 'nan':
                    dict_mh_code_by_name[mh_name.lower()] = mh_code

            # 矢量化全局映射计算
            p_chan_lower = df_pivot['供应商-渠道'].astype(str).str.strip().str.lower()
            p_prod_lower = df_pivot['买量产品'].astype(str).str.strip().str.lower()
            
            mapped_project = p_prod_lower.map(dict_project).fillna("")
            mapped_jindie = p_chan_lower.map(dict_supplier_jindie).fillna("")
            
            # 逐行填入并根据主体性质跨表精准提码
            for idx, row in df_pivot.iterrows():
                entity = str(row['核算主体']).upper().strip()
                
                # 豁免策略：NL主体免财务挂账，全部留空
                if entity == 'NL':
                    continue
                
                # 写入固定的项目编码与供应商名称
                df_pivot.at[idx, '项目'] = mapped_project.iloc[idx]
                current_jindie_name = mapped_jindie.iloc[idx]
                df_pivot.at[idx, '供应商-金蝶'] = current_jindie_name
                
                # 根据“主体”与“金蝶名称”，精确反查主数据对应编码，彻底断绝错位
                if current_jindie_name != "":
                    j_key = current_jindie_name.lower().strip()
                    if entity == 'CM':
                        df_pivot.at[idx, '供应商编码'] = dict_cm_code_by_name.get(j_key, "")
                    elif entity == 'MH':
                        df_pivot.at[idx, '供应商编码'] = dict_mh_code_by_name.get(j_key, "")
                        
        pivot_cols = ['买量产品', '核算主体', 'spent', '投放渠道', '开户方', '项目', '供应商-渠道', '供应商-金蝶', '供应商编码']
        df_pivot = df_pivot[pivot_cols]
        
        # ==========================================
        # 按钮一：原有常规业务分析总表 (数据对齐加固)
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
        for c in range(8, 8 + len(pivot_cols)):
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
        
        end_letter = openpyxl.utils.get_column_letter(7 + len(pivot_cols))
        ws_orig.merge_cells(f"H2:{end_letter}2")
        ws_orig["H2"] = "投放费用透气表"
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
                    if pivot_cols[c_idx-8] in ['项目', '供应商编码']:
                        cell.number_format = '@'

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
        # 按钮二：【给领导的汇总表】固定 6 页多 Sheet 纯流水
        # ==========================================
        wb_leader = openpyxl.Workbook()
        wb_leader.remove(wb_leader.active)
        
        font_l_hdr = Font(name="微软雅黑", size=10, bold=True)
        font_l_body = Font(name="微软雅黑", size=10)
        font_l_top_sub = Font(name="Arial", size=11, bold=True)
        font_l_grand_total = Font(name="Arial", size=11, bold=True, color="FF0000") 
        
        align_center = Alignment(horizontal="center", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        
        leader_thin_border = Border(
            left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC')
        )
        
        leader_headers = ["期间", "投放渠道", "开户服务商", "广告户名", "类别", "代投服务商", "消耗", "代投费", "投放待结算", "买量产品", "核算主体"]
        
        for sheet_name in REQUIRED_SHEETS.keys():
            ws_l = wb_leader.create_sheet(title=sheet_name)
            ws_l.views.sheetView[0].showGridLines = True
            
            df_sub = df_detail[df_detail['Target_Sheet'] == sheet_name] if 'Target_Sheet' in df_detail.columns else pd.DataFrame()
            
            if df_sub.empty:
                df_l_final = pd.DataFrame(columns=leader_headers)
                empty_row = {h: "" for h in leader_headers}
                empty_row['期间'] = current_month
                empty_row['买量产品'] = REQUIRED_SHEETS[sheet_name]['product']
                empty_row['核算主体'] = REQUIRED_SHEETS[sheet_name]['entity']
                df_l_final = pd.DataFrame([empty_row])
            else:
                df_l_final = df_sub[leader_headers].copy()
                df_l_final['期间'] = current_month
                
            data_end_row = max(4, len(df_l_final) + 3)
            
            if sheet_name == 'Advertising-Chapters':
                ws_l.cell(row=1, column=5, value="TTL:").font = font_l_hdr
                ws_l.cell(row=1, column=5).alignment = align_right
                
                ws_l.cell(row=1, column=6, value=f"=SUM('Advertising-Chapters:Advertising-RS N'!G4:G5000)").font = font_l_grand_total
                ws_l.cell(row=1, column=6).number_format = '#,##0.00'
                ws_l.cell(row=1, column=6).alignment = align_right
            
            ws_l.cell(row=1, column=7, value=f"=SUM(G4:G{data_end_row})").font = font_l_top_sub
            ws_l.cell(row=1, column=7).number_format = '#,##0.00'
            ws_l.cell(row=1, column=7).alignment = align_right
            
            ws_l.cell(row=1, column=9, value=f"=SUM(I4:I{data_end_row})").font = font_l_top_sub
            ws_l.cell(row=1, column=9).number_format = '#,##0.00'
            ws_l.cell(row=1, column=9).alignment = align_right
            
            for idx, h_name in enumerate(leader_headers, 1):
                cell = ws_l.cell(row=3, column=idx, value=h_name)
                cell.font = font_l_hdr
                cell.alignment = align_center
                cell.border = leader_thin_border
            ws_l.row_dimensions[3].height = 24
            
            for r_idx, row in enumerate(dataframe_to_rows(df_l_final, index=False, header=False), start=4):
                for c_idx, val in enumerate(row, start=1):
                    cell = ws_l.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = font_l_body
                    cell.border = leader_thin_border
                    
                    if leader_headers[c_idx-1] in ["消耗", "代投费", "投放待结算"]:
                        if val != "":
                            cell.number_format = '#,##0.00'
                        cell.alignment = align_right
                    else:
                        cell.alignment = align_center
                ws_l.row_dimensions[r_idx].height = 20
                
            for col in ws_l.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws_l.column_dimensions[col_letter].width = max(max_len + 3, 13)
                
        excel_data_leader = io.BytesIO()
        wb_leader.save(excel_data_leader)
        excel_data_leader.seek(0)
        
        # UI端按钮下载
        st.markdown("---")
        if not mp_file:
            st.warning("⚠️ 提示：您尚未上传 MP 数据映射表，右侧金蝶核算维度字段将暂时显示为空白。")
            
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"📊 2026年{month_label}投放费用计提表")
            st.download_button(
                label="点击下载新样式 Excel 报表",
                data=excel_data_orig,
                file_name=f"2026年{month_label}投放费用计提表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col2:
            st.success(f"👑 给领导的汇总表 (2026年{month_label}推广费用-各主体情况汇总.xlsx)")
            st.download_button(
                label="点击下载给领导的汇总表",
                data=excel_data_leader,
                file_name=f"2026年{month_label}推广费用-各主体情况汇总.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

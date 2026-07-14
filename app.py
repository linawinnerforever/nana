# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="金蝶云星空-费用重分类自动拆分工具", layout="wide")

st.title("📊 金蝶云星空 - 费用重分类自动生成标准模板工具")
st.markdown("""
本工具专门用于月底将核算维度为**公摊**的费用，按照指定的比例自动拆分到各个具体的**项目编码**中。
**终极修复版：全面兼容多工作表、智能对齐表头、强力数据去噪滤网！**
""")

st.sidebar.header("🛠️ 使用步骤")
st.sidebar.markdown("""
### 使用三步走：
1. **上传** 每月从金蝶导出的《费用明细表/TB余额表》
2. **上传** 后台维护的《部门项目分摊比例表》
3. **上传** 金蝶标准的《引入凭证Excel模版》
4. **点击** 生成按钮，下载即可直接导入金蝶
""")

# 创建上传区域
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. 上传费用明细表 (TB)")
    source_file = st.file_uploader("支持 .xlsx 格式，需含：科目编码、科目名称、核算维度编码、核算维度名称、待拆分金额", type=["xlsx"], key="src")

with col2:
    st.subheader("2. 上传部门项目分摊比例表")
    ratio_file = st.file_uploader("支持 .xlsx 格式，需含：成本中心编号，以及横向的项目列(如001, 002等)", type=["xlsx"], key="ratio")

with col3:
    st.subheader("3. 上传金蝶引入凭证空模版")
    template_file = st.file_uploader("支持 .xlsx 格式的官方标准凭证导入模板", type=["xlsx"], key="template")

if source_file and ratio_file and template_file:
    try:
        # 智能全Sheet扫描：防止读错非数据工作表
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
        
        ratio_excel = pd.ExcelFile(ratio_file)
        valid_ratio_sheet = ratio_excel.sheet_names[0]
        ratio_header_idx = 0
        for sheet in ratio_excel.sheet_names:
            df_check = pd.read_excel(ratio_file, sheet_name=sheet, header=None)
            for idx, row in df_check.iterrows():
                if "成本中心编号" in row.values:
                    valid_ratio_sheet = sheet
                    ratio_header_idx = idx
                    break
                    
        df_ratio = pd.read_excel(ratio_file, sheet_name=valid_ratio_sheet, skiprows=ratio_header_idx)
        df_ratio.columns = df_ratio.columns.astype(str).str.strip()
        
        df_template = pd.read_excel(template_file, header=None)
        
        st.success(f"✅ 数据加载成功！已锁定费用表工作表 [{valid_src_sheet}] 和比例表工作表 [{valid_ratio_sheet}]")
        
        with st.expander("🔍 预览成功识别到的有效表头"):
            st.markdown("**费用明细表 (TB) 表头：**")
            st.write(list(df_source.columns))
            st.markdown("**拆分比例库表头：**")
            st.write(list(df_ratio.columns))
            
        if st.button("🚀 开始自动化重分类并生成金蝶模版"):
            fixed_cols = ['成本中心编号', '成本中心名称', '大部门分类', '项目', '分摊逻辑', '分摊过渡部门', '分摊类型', '合计']
            project_cols = [col for col in df_ratio.columns if col not in fixed_cols and not col.startswith('Unnamed:')]
            
            # 强力去噪：强制转化数字并清除中文干扰行
            df_source['待拆分金额_numeric'] = pd.to_numeric(df_source['待拆分金额'], errors='coerce')
            df_to_split = df_source[df_source['待拆分金额_numeric'].notna() & (df_source['待拆分金额_numeric'] != 0)]
            
            if df_to_split.empty:
                st.warning("⚠️ 费用明细表中未发现‘待拆分金额’不为0的有效数字数据行，请确认表格中该列是否填有数字。")
            
            tech_headers = df_template.iloc[0].tolist()
            cn_headers = df_template.iloc[1].tolist()
            
            new_rows = []
            error_logs = []
            
            base_info = {
                'FAccountBookID': '001',
                'FAccountBookID#Name': '北京枫悦互动科技有限公司',
                'FDate': '2026-06-30',
                'FBUSDATE': '2026-06-30',
                'FYEAR': 2026,
                'FPERIOD': 6,
                'FVOUCHERGROUPID': '記',
                'FVOUCHERGROUPID#Name': '记账凭证',
                'FVOUCHERGROUPNO': '1',
                'FATTACHMENTS': 0,
                'FISADJUSTVOUCHER': '否',
                'FACCBOOKORGID': '001',
                'FACCBOOKORGID#Name': '北京枫悦互动科技有限公司',
                'FCURRENCYID': 'PRE001',
                'FCURRENCYID#Name': '人民币',
                'FEXCHANGERATETYPE': 'HLTX01_SYS',
                'FEXCHANGERATETYPE#Name': '固定汇率',
                'FEXCHANGERATE': 1
            }
            
            if df_template.shape[0] > 2:
                for i, col_name in enumerate(df_template.iloc[0]):
                    if pd.notna(col_name) and str(col_name) in base_info:
                        val = df_template.iloc[2, i]
                        if pd.notna(val):
                            base_info[str(col_name)] = val
            
            entry_idx = 1
            
            for idx, row in df_to_split.iterrows():
                try:
                    sub_code = str(row['科目编码']).strip()
                    sub_name = str(row['科目名称']).strip()
                    dim_code = str(row['核算维度编码']).strip()
                    
                    if sub_code == '科目编码' or sub_code == 'nan':
                        continue
                        
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
                        error_logs.append(f"科目 {sub_code}: 核算维度编码 '{dim_code}' 在比例表中找不到对应的【成本中心编号】，已跳过。")
                        continue
                        
                    ratio_row_data = df_ratio[df_ratio['成本中心编号'].astype(str).str.strip() == cc_code].iloc[0]
                    
                    valid_projects = []
                    for proj in project_cols:
                        val = ratio_row_data[proj]
                        try:
                            val_float = float(val)
                            if val_float > 0:
                                valid_projects.append({'proj_code': proj, 'ratio': val_float})
                        except:
                            pass
                            
                    if not valid_projects:
                        error_logs.append(f"科目 {sub_code}: 成本中心 '{cc_code}' 未配置有效分摊比例，已跳过。")
                        continue
                    
                    df_valid_proj = pd.DataFrame(valid_projects).sort_values(by='ratio', ascending=False)
                    
                    # 1. 借方负数行 (冲销原公摊)
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
                    
                    # 2. 借方正数行 (分摊至具体项目)
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
                            if k in tech_headers: pos_row[tech_headers.index(k)] = v
                        
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
                        
                except Exception as e:
                    error_logs.append(f"处理行发生错误: {str(e)}")
            
            if error_logs:
                with st.sidebar.expander("⚠️ 处理日志与提示"):
                    for log in error_logs:
                        st.warning(log)
            
            if new_rows:
                final_df = pd.DataFrame([tech_headers, cn_headers] + new_rows)
                st.success("🎉 金蝶标准凭证上传模版数据已自动填充并完美排序！")
                st.markdown("**金蝶上传模板填充数据预览：**")
                st.dataframe(final_df.iloc[2:15])
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 点击下载已自动填充好的金蝶上传凭证Excel文件",
                    data=processed_data,
                    file_name="金蝶云星空重分类引入凭证(已填充).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("❌ 未能生成任何凭证行，请确认数据是否匹配。")
                
    except Exception as e:
        st.error(f"解析文件或匹配金蝶模版失败。错误信息: {e}")
else:
    st.info("💡 请在上方上传您的三张 Excel 原始表，工具会自动激活计算。")

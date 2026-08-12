import pandas as pd
import numpy as np
import datetime
import calendar

def build_kingdee_voucher(draft_filepath, output_filepath, date_str="2026-07-31"):
    """
    将拆分底稿自动转换为金蝶上传凭证模板
    :param draft_filepath: 输入的拆分底稿 Excel 路径
    :param output_filepath: 输出的金蝶凭证 Excel 路径
    :param date_str: 凭证日期，格式 YYYY-MM-DD
    """
    # 1. 解析日期参数
    voucher_date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    year = voucher_date.year
    month = voucher_date.month
    
    # 获取当月最后一天
    _, last_day = calendar.monthrange(year, month)
    date_formatted = f"{year}-{month:02d}-{last_day:02d}"
    explanation = f"计提{year}年{month}月Claude消耗-主营业务成本_软件服务费"

    # 2. 读取底稿原始数据
    df_raw = pd.read_excel(draft_filepath, header=None)
    
    # 获取项目段编码行（第0行）及对应数据列索引
    # 根据样本底稿，数据列通常从第6列开始（001, 002, 003...）
    project_segment_row = df_raw.iloc[0].values
    
    # 识别列头中有有效项目编码（如001, 002, 010等）的列
    data_cols = []
    for col_idx in range(len(project_segment_row)):
        val = str(project_segment_row[col_idx]).strip()
        if val != 'nan' and val != '' and val.isdigit():
            data_cols.append((col_idx, val.zfill(3))) # 补齐3位编码，如 '001'
            
    # 查找成本中心数据行（成本中心编码在第1列，从第4行开始）
    rows_data = []
    for r in range(3, len(df_raw)):
        dept_code = df_raw.iloc[r, 1]
        if pd.isna(dept_code):
            continue
        # 转为字符串并清洗
        dept_code_str = str(int(dept_code)) if isinstance(dept_code, (int, float)) else str(dept_code).strip()
        
        # 遍历每个项目段列获取金额
        for col_idx, proj_code in data_cols:
            amount = df_raw.iloc[r, col_idx]
            # 过滤 NaN、空值及 0（保留 0.01 或非零值）
            if pd.notna(amount):
                try:
                    amount_val = round(float(amount), 2)
                    if amount_val != 0:
                        rows_data.append({
                            'dept_code': dept_code_str,
                            'proj_code': proj_code,
                            'amount': amount_val
                        })
                except ValueError:
                    continue

    # 3. 构造金蝶凭证表头与结构 (列头包含双行标题)
    header_row0 = [
        'FBillHead(GL_VOUCHER)', 'FAccountBookID', 'FAccountBookID#Name', 'FDate', 'FBUSDATE', 'FYEAR', 'FPERIOD',
        'FVOUCHERGROUPID', 'FVOUCHERGROUPID#Name', 'FVOUCHERGROUPNO', 'FATTACHMENTS', 'FISADJUSTVOUCHER',
        'FACCBOOKORGID', 'FACCBOOKORGID#Name', 'FSourceBillKey', 'FSourceBillKey#Name', 'FIMPORTVERSION',
        '*Split*1', 'FEntity', 'FEXPLANATION', 'FACCOUNTID', 'FACCOUNTID#Name', 'FDetailID#FF100003',
        'FDetailID#FF100003#Name', 'FDetailID#FF100002', 'FDetailID#FF100002#Name', 'FDetailID#FFLEX16',
        'FDetailID#FFLEX16#Name', 'FDetailID#FFLEX15', 'FDetailID#FFLEX15#Name', 'FDetailID#FFLEX14',
        'FDetailID#FFLEX14#Name', 'FDetailID#FFLEX13', 'FDetailID#FFLEX13#Name', 'FDetailID#FFLEX12',
        'FDetailID#FFLEX12#Name', 'FDetailID#FFLEX11', 'FDetailID#FFLEX11#Name', 'FDetailID#FFlex10',
        'FDetailID#FFlex10#Name', 'FDetailID#FFLEX9', 'FDetailID#FFLEX9#Name', 'FDetailID#FFlex8',
        'FDetailID#FFlex8#Name', 'FDetailID#FFlex7', 'FDetailID#FFlex7#Name', 'FDetailID#FFlex6',
        'FDetailID#FFlex6#Name', 'FDetailID#FFlex5', 'FDetailID#FFlex5#Name', 'FDetailID#FFlex4',
        'FDetailID#FFlex4#Name', 'FDetailID#FF100004', 'FDetailID#FF100004#Name', 'FDetailID#FF100005',
        'FDetailID#FF100005#Name', 'FCURRENCYID', 'FCURRENCYID#Name', 'FEXCHANGERATETYPE', 'FEXCHANGERATETYPE#Name',
        'FEXCHANGERATE', 'FUnitId', 'FUnitId#Name', 'FPrice', 'FQty', 'FAMOUNTFOR', 'FDEBIT', 'FCREDIT',
        'FSettleTypeID', 'FSettleTypeID#Name', 'FSETTLENO', 'FBUSNO', 'FEXPORTENTRYID'
    ]
    
    header_row1 = [
        '*单据头(序号)', '*(单据头)账簿#编码', '(单据头)账簿#名称', '*(单据头)日期', '(单据头)业务日期', '(单据头)会计年度',
        '(单据头)期间', '*(单据头)凭证字#编码', '(单据头)凭证字#名称', '*(单据头)凭证号', '(单据头)附件数', '(单据头)是否调整期凭证',
        '*(单据头)核算组织#编码', '(单据头)核算组织#名称', '(单据头)业务类型#编码', '(单据头)业务类型#名称', '(单据头)引入版本号',
        '间隔列', '*分录(序号)', '(分录)摘要', '*(分录)科目编码#编码', '(分录)科目编码#名称', '(分录)北京公司项目名#编码',
        '(分录)北京公司项目名#名称(Null)', '(分录)项目段#编码', '(分录)项目段#名称(Null)', '(分录)其他往来单位#编码',
        '(分录)其他往来单位#名称(Null)', '(分录)银行账号#编码', '(分录)银行账号#名称(Null)', '(分录)银行#编码',
        '(分录)银行#名称(Null)', '(分录)客户分组#编码', '(分录)客户分组#名称(Null)', '(分录)物料分组#编码',
        '(分录)物料分组#名称(Null)', '(分录)组织机构#编码', '(分录)组织机构#名称(Null)', '(分录)资产类别#编码',
        '(分录)资产类别#名称(Null)', '(分录)费用项目#编码', '(分录)费用项目#名称(Null)', '(分录)物料#编码',
        '(分录)物料#名称(Null)', '(分录)员工#编码', '(分录)员工#名称(Null)', '(分录)客户#编码',
        '(分录)客户#名称(Null)', '(分录)部门#编码', '(分录)部门#名称(Null)', '(分录)供应商#编码',
        '(分录)供应商#名称(Null)', '(分录)NL剧集#编码', '(分录)NL剧集#名称(Null)', '(分录)海南剧集#编码',
        '(分录)海南剧集#名称(Null)', '*(分录)币别#编码', '(分录)币别#名称', '*(分录)汇率类型#编码',
        '(分录)汇率类型#名称', '(分录)汇率', '(分录)单位#编码', '(分录)单位#名称', '(分录)单价', '(分录)数量',
        '(分录)原币金额', '(分录)借方金额', '(分录)贷方金额', '(分录)结算方式#编码', '(分录)结算方式#名称', '(分录)结算号',
        '(分录)业务编号', '(分录)现金流量#分录ID'
    ]

    result_rows = [header_row0, header_row1]
    
    total_debit_amount = 0.0
    entry_seq = 1

    # 4. 生成借方分录
    for item in rows_data:
        amt = item['amount']
        total_debit_amount += amt
        
        is_first = (entry_seq == 1)
        row = [
            1 if is_first else np.nan,              # *单据头(序号)
            '002' if is_first else np.nan,          # 账簿编码
            np.nan,                                 # 账簿名称
            date_formatted if is_first else np.nan, # 日期
            date_formatted if is_first else np.nan, # 业务日期
            year if is_first else np.nan,           # 会计年度
            month if is_first else np.nan,          # 期间
            'PRE001' if is_first else np.nan,       # 凭证字
            np.nan,
            1 if is_first else np.nan,              # 凭证号
            np.nan, np.nan,
            100 if is_first else np.nan,            # 核算组织
            np.nan, np.nan, np.nan, np.nan, np.nan,
            entry_seq,                              # *分录(序号)
            explanation,                            # 摘要
            '6401.21',                              # 借方科目编码
            np.nan, np.nan, np.nan,
            item['proj_code'],                      # 项目段编码
            np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
            np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
            item['dept_code'],                      # 部门编码
            np.nan,
            'VEN02027',                             # 供应商编码
            np.nan, np.nan, np.nan, np.nan, np.nan,
            'PRE007',                               # 币别
            '美元',
            'HLTX01_SYS',                           # 汇率类型
            '固定汇率',
            1,                                      # 汇率
            np.nan, np.nan, np.nan, np.nan,
            amt,                                    # 原币金额
            amt,                                    # 借方金额
            np.nan,                                 # 贷方金额
            np.nan, np.nan, np.nan, np.nan, np.nan
        ]
        result_rows.append(row)
        entry_seq += 1

    # 5. 生成贷方分录 (应付账款 2202.01)
    total_debit_amount = round(total_debit_amount, 2)
    credit_row = [
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        entry_seq,                                  # 分录序号
        explanation,                                # 摘要
        '2202.01',                                  # 贷方科目
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
        'VEN02027',                                 # 供应商编码
        np.nan, np.nan, np.nan, np.nan, np.nan,
        'PRE007', '美元', 'HLTX01_SYS', '固定汇率', 1,
        np.nan, np.nan, np.nan, np.nan,
        np.nan,                                     # 原币金额
        np.nan,                                     # 借方金额
        total_debit_amount,                         # 贷方金额
        np.nan, np.nan, np.nan, np.nan, np.nan
    ]
    result_rows.append(credit_row)

    # 6. 保存导出为 Excel
    out_df = pd.DataFrame(result_rows)
    out_df.to_excel(output_filepath, index=False, header=False)
    print(f"成功生成金蝶凭证上传文件：{output_filepath}，共 {entry_seq} 条分录，总金额：{total_debit_amount}")

# 示例运行测试
if __name__ == '__main__':
    draft_file = '测试-样本Claude拆分底稿-202607.xlsx'
    output_file = '生成的金蝶上传凭证-202607.xlsx'
    build_kingdee_voucher(draft_file, output_file, date_str="2026-07-31")

import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string as column_letter_to_number, get_column_letter as column_number_to_letter
from datetime import datetime

# 为了重用之前的复杂代码，直接引入旧的 DataProcessor (作为示例平滑迁移，实际可全盘重构进这里)
from core_engine.data_processor import DataProcessor

class BankPlugin:
    @staticmethod
    def bank_write_data(context, params):
        print(f"开始处理银行数据")
        ws = context['ws']
        reader = context['data_reader']
        sheet_config = context['sheet_config']
        source_sheet = sheet_config['source_sheet']
        indicator_col_map = sheet_config['indicators']
        start_row = params.get('start_row', 50)
        start_column = params.get('start_column', 2)
        start_date = params.get('start_date', '2021-01-01')
        date_format = params.get('date_format', 'yyyy-mm')
        start_date_ts = pd.to_datetime(start_date)

        # 1. 从缓存中获取所需所有指标数据
        indicator_dfs = {}
        for indicator in indicator_col_map:
            ind_data = reader.read_indicator_data(source_sheet, indicator)
            df = ind_data["data"].copy()

            if df.empty or '日期' not in df.columns:
                indicator_dfs[indicator] = pd.DataFrame(columns=['日期'])
                print(f"    - 工作表[{source_sheet}] 未找到指标 {indicator}，该列跳过写入")
                continue

            df['日期'] = pd.to_datetime(df['日期'], errors='coerce')

            # 只保留 start_date 及以后数据
            df = df[df['日期'].notna() & (df['日期'] >= start_date_ts)].copy()
            df = df.sort_values('日期', ascending=False).reset_index(drop=True)

            indicator_dfs[indicator] = df

        # 2. 使用第一个指标的数据作为基础日期表，并筛选季度数据（3、6、9、12月）
        first_indicator = list(indicator_col_map.keys())[0]
        first_df = indicator_dfs[first_indicator].copy()

        if first_df.empty:
            print(f"    - {source_sheet} 的第一个指标 {first_indicator} 无有效数据，跳过写入")
            return

        # 筛选只包含3、6、9、12月的数据
        quarterly_df = first_df[first_df['日期'].dt.month.isin([3, 6, 9, 12])].copy()
        
        if quarterly_df.empty:
            print(f"    - {source_sheet} 的第一个指标 {first_indicator} 中没有3、6、9、12月的数据，跳过写入")
            return

        base_df = quarterly_df.reset_index(drop=True)

        # 3. 写入统一日期列和指标数据（应用单位转换）
        for i, row in base_df.iterrows():
            current_row = start_row + i
            current_date = row['日期']

            # 第1列写日期
            date_cell = ws.cell(row=current_row, column=start_column, value=current_date)
            date_cell.number_format = date_format

        # 4. 写入指标数据，全部按 base_df 的日期对齐（应用单位转换）
        base_date_list = base_df['日期'].tolist()

        for indicator, col in list(indicator_col_map.items()):
            df = indicator_dfs[indicator].copy()
            # 对每个指标也筛选季度数据
            df_quarterly = df[df['日期'].dt.month.isin([3, 6, 9, 12])].copy()
            
            if df_quarterly.empty:
                print(f"    - 指标 {indicator} 中没有3、6、9、12月的数据")
                continue
                
            value_col = df_quarterly.columns[-1]
            value_map = dict(zip(df_quarterly['日期'], df_quarterly[value_col]))
            col_num = column_letter_to_number(col) if isinstance(col, str) else int(col)    
            for i, date in enumerate(base_date_list):
                value = value_map.get(date, None)
                ws.cell(
                    row=start_row + i,
                    column=col_num,
                    value=value
                )

        print(f"    - 完成银行基础数据写入")
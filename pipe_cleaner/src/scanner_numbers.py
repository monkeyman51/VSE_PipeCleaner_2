import pandas as pd

cover_page_df = pd.read_excel(f'../input/commodity_numbers.xlsx', sheet_name='All Gen')
cover_page_df.to_csv('commodity_numbers.csv', index=False)

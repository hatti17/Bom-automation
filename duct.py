import pandas as pd
from openpyxl import load_workbook

excel_file = 'Documents/Input/duct.xlsx'
df = pd.read_excel(excel_file)

# *********************** carriageway_S_tarmac loc no and wayleave no  ************************************************

footway_s_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Footway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Access')]

footway_s_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Footway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Access')
                         &  (df['96mm'] == '2x96')]

num_of_rows_fw_s_tarmac = footway_s_tarmac['length'].sum()

num_of_rows_fw_s_tarmac_2x96 = footway_s_tarmac_2x96['length'].sum()

num_of_rows_fw_s_tarmac_total = num_of_rows_fw_s_tarmac + num_of_rows_fw_s_tarmac_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["E54"] = num_of_rows_fw_s_tarmac_total
wb.save(filename)


import pandas as pd
import numpy as np
from openpyxl import load_workbook

excel_file = 'Documents/Input/duct.xlsx'
df = pd.read_excel(excel_file)

# *********************** carriageway_S+T_modular loc no and wayleave no  *********************************************

carriageway_s_t_modular = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular' ) & (df['type'] == 'Access & Trunk') ]

carriageway_s_t_modular_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular') & (df['type'] == 'Access & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_t_modular = carriageway_s_t_modular['length'].sum()

num_of_rows_cw_s_t_modular_2x96 = carriageway_s_t_modular_2x96['length'].sum()

num_of_rows_cw_s_t_modular_total = num_of_rows_cw_s_t_modular + num_of_rows_cw_s_t_modular_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["B38"] = num_of_rows_cw_s_t_modular_total
wb.save(filename)

# *********************** carriageway_S+T_concrete loc no and wayleave no  *********************************************

carriageway_s_t_concrete = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Access & Trunk') ]

carriageway_s_t_concrete_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Access & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_t_concrete = carriageway_s_t_concrete['length'].sum()

num_of_rows_cw_s_t_concrete_2x96 = carriageway_s_t_concrete_2x96['length'].sum()

num_of_rows_cw_s_t_concrete_total = num_of_rows_cw_s_t_concrete + num_of_rows_cw_s_t_concrete_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["C38"] = num_of_rows_cw_s_t_concrete_total
wb.save(filename)

# *********************** carriageway_S+T_unmade and grassverge loc no and wayleave no  *********************************************

carriageway_s_t_grassverge = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Access & Trunk')]

carriageway_s_t_unmade = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Access & Trunk')]

carriageway_s_t_grassverge_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Access & Trunk') & (df['96mm'] == '2x96')]

carriageway_s_t_unmade_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Access & Trunk') & (df['96mm'] == '2x96')]


num_of_rows_cw_s_t_grassverge = carriageway_s_t_grassverge['length'].sum()

num_of_rows_cw_s_t_unmade = carriageway_s_t_unmade['length'].sum()

num_of_rows_cw_s_t_grassverge_2x96 = carriageway_s_t_grassverge_2x96['length'].sum()

num_of_rows_cw_s_t_unmade_2x96 = carriageway_s_t_unmade_2x96['length'].sum()

num_of_rows_cw_s_t_unmade_grassverge_total = num_of_rows_cw_s_t_grassverge + num_of_rows_cw_s_t_unmade + num_of_rows_cw_s_t_grassverge_2x96 + num_of_rows_cw_s_t_unmade_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["D38"] = num_of_rows_cw_s_t_unmade_grassverge_total
wb.save(filename)

# *********************** carriageway_S+T_tarmac loc no and wayleave no  *********************************************

carriageway_s_t_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac' ) & (df['type'] == 'Access & Trunk') ]

carriageway_s_t_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Access & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_t_tarmac = carriageway_s_t_tarmac['length'].sum()

num_of_rows_cw_s_t_tarmac_2x96 = carriageway_s_t_tarmac_2x96['length'].sum()

num_of_rows_cw_s_t_tarmac_total = num_of_rows_cw_s_t_tarmac + num_of_rows_cw_s_t_tarmac_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["E38"] = num_of_rows_cw_s_t_tarmac_total
wb.save(filename)

# *********************** carriageway_D+T_modular loc no and wayleave no  *********************************************

carriageway_d_t_modular = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular' ) & (df['type'] == 'Distribution & Trunk') ]

carriageway_d_t_modular_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular') & (df['type'] == 'Distribution & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_d_t_modular = carriageway_d_t_modular['length'].sum()

num_of_rows_cw_d_t_modular_2x96 = carriageway_d_t_modular_2x96['length'].sum()

num_of_rows_cw_d_t_modular_total = num_of_rows_cw_d_t_modular + num_of_rows_cw_d_t_modular_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["B39"] = num_of_rows_cw_d_t_modular_total
wb.save(filename)

# *********************** carriageway_D+T_concrete loc no and wayleave no  *********************************************

carriageway_d_t_concrete = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Distribution & Trunk') ]

carriageway_d_t_concrete_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Distribution & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_d_t_concrete = carriageway_d_t_concrete['length'].sum()

num_of_rows_cw_d_t_concrete_2x96 = carriageway_d_t_concrete_2x96['length'].sum()

num_of_rows_cw_d_t_concrete_total = num_of_rows_cw_d_t_concrete + num_of_rows_cw_d_t_concrete_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["C39"] = num_of_rows_cw_d_t_concrete_total
wb.save(filename)

# *********************** carriageway_D+T_unmade and grassverge loc no and wayleave no  *********************************************

carriageway_d_t_grassverge = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Distribution & Trunk')]

carriageway_d_t_unmade = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Distribution & Trunk')]

carriageway_d_t_grassverge_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Distribution & Trunk')
                         & (df['96mm'] == '2x96')]

carriageway_d_t_unmade_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Distribution & Trunk')
                         & (df['96mm'] == '2x96')]


num_of_rows_cw_d_t_grassverge = carriageway_d_t_grassverge['length'].sum()

num_of_rows_cw_d_t_unmade = carriageway_d_t_unmade['length'].sum()

num_of_rows_cw_d_t_grassverge_2x96 = carriageway_d_t_grassverge_2x96['length'].sum()

num_of_rows_cw_d_t_unmade_2x96 = carriageway_d_t_unmade_2x96['length'].sum()

num_of_rows_cw_d_t_unmade_grassverge_total = num_of_rows_cw_d_t_grassverge + num_of_rows_cw_d_t_unmade + num_of_rows_cw_d_t_grassverge_2x96 + num_of_rows_cw_d_t_unmade_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["D39"] = num_of_rows_cw_d_t_unmade_grassverge_total
wb.save(filename)

# *********************** carriageway_D+T_tarmac loc no and wayleave no  *********************************************

carriageway_d_t_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac' ) & (df['type'] == 'Distribution & Trunk') ]

carriageway_d_t_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Distribution & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_d_t_tarmac = carriageway_d_t_tarmac['length'].sum()

num_of_rows_cw_d_t_tarmac_2x96 = carriageway_d_t_tarmac_2x96['length'].sum()

num_of_rows_cw_d_t_tarmac_total = num_of_rows_cw_d_t_tarmac + num_of_rows_cw_d_t_tarmac_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["E39"] = num_of_rows_cw_d_t_tarmac_total
wb.save(filename)

# *********************** carriageway_D_modular loc no and wayleave no  *********************************************

carriageway_d_modular = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular' ) & (df['type'] == 'Distribution') ]

carriageway_d_modular_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular') & (df['type'] == 'Distribution')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_d_modular = carriageway_d_modular['length'].sum()

num_of_rows_cw_d_modular_2x96 = carriageway_d_modular_2x96['length'].sum()

num_of_rows_cw_d_modular_total = num_of_rows_cw_d_modular + num_of_rows_cw_d_modular_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["B40"] = num_of_rows_cw_d_modular_total
wb.save(filename)

# *********************** carriageway_D_concrete loc no and wayleave no  *********************************************

carriageway_d_concrete = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Distribution') ]

carriageway_d_concrete_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Distribution')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_d_concrete = carriageway_d_concrete['length'].sum()

num_of_rows_cw_d_concrete_2x96 = carriageway_d_concrete_2x96['length'].sum()

num_of_rows_cw_d_concrete_total = num_of_rows_cw_d_concrete + num_of_rows_cw_d_concrete_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["C40"] = num_of_rows_cw_d_concrete_total
wb.save(filename)

# *********************** carriageway_D_unmade and grassverge loc no and wayleave no  *********************************************

carriageway_d_grassverge = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Distribution')]

carriageway_d_unmade = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Distribution')]

carriageway_d_grassverge_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Distribution')
                         & (df['96mm'] == '2x96')]

carriageway_d_unmade_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Distribution')
                         & (df['96mm'] == '2x96')]


num_of_rows_cw_d_grassverge = carriageway_d_grassverge['length'].sum()

num_of_rows_cw_d_unmade = carriageway_d_unmade['length'].sum()

num_of_rows_cw_d_grassverge_2x96 = carriageway_d_grassverge_2x96['length'].sum()

num_of_rows_cw_d_unmade_2x96 = carriageway_d_unmade_2x96['length'].sum()

num_of_rows_cw_d_unmade_grassverge_total = num_of_rows_cw_d_grassverge + num_of_rows_cw_d_unmade + num_of_rows_cw_d_grassverge_2x96 + num_of_rows_cw_d_unmade_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["D40"] = num_of_rows_cw_d_unmade_grassverge_total
wb.save(filename)

# *********************** carriageway_D_tarmac loc no and wayleave no  *********************************************

carriageway_d_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac' ) & (df['type'] == 'Distribution') ]

carriageway_d_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Distribution')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_d_tarmac = carriageway_d_tarmac['length'].sum()

num_of_rows_cw_d_tarmac_2x96 = carriageway_d_tarmac_2x96['length'].sum()

num_of_rows_cw_d_tarmac_total = num_of_rows_cw_d_tarmac + num_of_rows_cw_d_tarmac_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["E40"] = num_of_rows_cw_d_tarmac_total
wb.save(filename)

# *********************** footway_S_tarmac loc no and wayleave no  ************************************************

footway_s_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Footway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Access') ]

footway_s_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Footway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Access')
                         &  (df['96mm'] == '2x96') ]

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


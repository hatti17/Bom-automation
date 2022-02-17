import pandas as pd
import numpy as np
from openpyxl import load_workbook

excel_file = 'Documents/Input/duct.xlsx'
df = pd.read_excel(excel_file)

#carriageway loc no and wayleave no
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

# *********************** carriageway_S+D+T_modular loc no and wayleave no  *********************************************

carriageway_s_d_t_modular = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular' ) & (df['type'] == 'Acces, Distribution & Trunk') ]

carriageway_s_d_t_modular_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular') & (df['type'] == 'Acces, Distribution & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_d_t_modular = carriageway_s_d_t_modular['length'].sum()

num_of_rows_cw_s_d_t_modular_2x96 = carriageway_s_d_t_modular_2x96['length'].sum()

num_of_rows_cw_s_d_t_modular_total = num_of_rows_cw_s_d_t_modular + num_of_rows_cw_s_d_t_modular_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["B41"] = num_of_rows_cw_s_d_t_modular_total
wb.save(filename)

# *********************** carriageway_S+D+T_concrete loc no and wayleave no  *********************************************

carriageway_s_d_t_concrete = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Acces, Distribution & Trunk') ]

carriageway_s_d_t_concrete_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Acces, Distribution & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_d_t_concrete = carriageway_s_d_t_concrete['length'].sum()

num_of_rows_cw_s_d_t_concrete_2x96 = carriageway_s_d_t_concrete_2x96['length'].sum()

num_of_rows_cw_s_d_t_concrete_total = num_of_rows_cw_s_d_t_concrete + num_of_rows_cw_s_d_t_concrete_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["C41"] = num_of_rows_cw_s_d_t_concrete_total
wb.save(filename)

# *********************** carriageway_S+D+T_unmade and grassverge loc no and wayleave no  *********************************************

carriageway_s_d_t_grassverge = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Acces, Distribution & Trunk')]

carriageway_s_d_t_unmade = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Acces, Distribution & Trunk')]

carriageway_s_d_t_grassverge_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Acces, Distribution & Trunk')
                         & (df['96mm'] == '2x96')]

carriageway_s_d_t_unmade_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Acces, Distribution & Trunk')
                         & (df['96mm'] == '2x96')]


num_of_rows_cw_s_d_t_grassverge = carriageway_s_d_t_grassverge['length'].sum()

num_of_rows_cw_s_d_t_unmade = carriageway_s_d_t_unmade['length'].sum()

num_of_rows_cw_s_d_t_grassverge_2x96 = carriageway_s_d_t_grassverge_2x96['length'].sum()

num_of_rows_cw_s_d_t_unmade_2x96 = carriageway_s_d_t_unmade_2x96['length'].sum()

num_of_rows_cw_s_d_t_unmade_grassverge_total = num_of_rows_cw_s_d_t_grassverge + num_of_rows_cw_s_d_t_unmade + num_of_rows_cw_s_d_t_grassverge_2x96 + num_of_rows_cw_s_d_t_unmade_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["D41"] = num_of_rows_cw_s_d_t_unmade_grassverge_total
wb.save(filename)

# *********************** carriageway_S+D+T_tarmac loc no and wayleave no  *********************************************

carriageway_s_d_t_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac' ) & (df['type'] == 'Acces, Distribution & Trunk') ]

carriageway_s_d_t_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Acces, Distribution & Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_d_t_tarmac = carriageway_s_d_t_tarmac['length'].sum()

num_of_rows_cw_s_d_t_tarmac_2x96 = carriageway_s_d_t_tarmac_2x96['length'].sum()

num_of_rows_cw_s_d_t_tarmac_total = num_of_rows_cw_s_d_t_tarmac + num_of_rows_cw_s_d_t_tarmac_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["E41"] = num_of_rows_cw_s_d_t_tarmac_total
wb.save(filename)

# *********************** carriageway_S+D_modular loc no and wayleave no  *********************************************

carriageway_s_d_modular = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular' ) & (df['type'] == 'Access & Distribution') ]

carriageway_s_d_modular_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular') & (df['type'] == 'Access & Distribution')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_d_modular = carriageway_s_d_modular['length'].sum()

num_of_rows_cw_s_d_modular_2x96 = carriageway_s_d_modular_2x96['length'].sum()

num_of_rows_cw_s_d_modular_total = num_of_rows_cw_s_d_modular + num_of_rows_cw_s_d_modular_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["B42"] = num_of_rows_cw_s_d_modular_total
wb.save(filename)

# *********************** carriageway_S+D_concrete loc no and wayleave no  *********************************************

carriageway_s_d_concrete = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Access & Distribution') ]

carriageway_s_d_concrete_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Access & Distribution')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_d_concrete = carriageway_s_d_concrete['length'].sum()

num_of_rows_cw_s_d_concrete_2x96 = carriageway_s_d_concrete_2x96['length'].sum()

num_of_rows_cw_s_d_concrete_total = num_of_rows_cw_s_d_concrete + num_of_rows_cw_s_d_concrete_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["C42"] = num_of_rows_cw_s_d_concrete_total
wb.save(filename)

# *********************** carriageway_S+D_unmade and grassverge loc no and wayleave no  *********************************************

carriageway_s_d_grassverge = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Access & Distribution')]

carriageway_s_d_unmade = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Access & Distribution')]

carriageway_s_d_grassverge_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Access & Distribution')
                         & (df['96mm'] == '2x96')]

carriageway_s_d_unmade_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Access & Distribution')
                         & (df['96mm'] == '2x96')]


num_of_rows_cw_s_d_grassverge = carriageway_s_d_grassverge['length'].sum()

num_of_rows_cw_s_d_unmade = carriageway_s_d_unmade['length'].sum()

num_of_rows_cw_s_d_grassverge_2x96 = carriageway_s_d_grassverge_2x96['length'].sum()

num_of_rows_cw_s_d_unmade_2x96 = carriageway_s_d_unmade_2x96['length'].sum()

num_of_rows_cw_s_d_unmade_grassverge_total = num_of_rows_cw_s_d_grassverge + num_of_rows_cw_s_d_unmade + num_of_rows_cw_s_d_grassverge_2x96 + num_of_rows_cw_s_d_unmade_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["D42"] = num_of_rows_cw_s_d_unmade_grassverge_total
wb.save(filename)

# *********************** carriageway_S+D_tarmac loc no and wayleave no  *********************************************

carriageway_s_d_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac' ) & (df['type'] == 'Access & Distribution') ]

carriageway_s_d_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Access & Distribution')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_d_tarmac = carriageway_s_d_tarmac['length'].sum()

num_of_rows_cw_s_d_tarmac_2x96 = carriageway_s_d_tarmac_2x96['length'].sum()

num_of_rows_cw_s_d_tarmac_total = num_of_rows_cw_s_d_tarmac + num_of_rows_cw_s_d_tarmac_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["E42"] = num_of_rows_cw_s_d_tarmac_total
wb.save(filename)

# *********************** carriageway_S_modular loc no and wayleave no  *********************************************

carriageway_s_modular = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular' ) & (df['type'] == 'Access') ]

carriageway_s_modular_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular') & (df['type'] == 'Access')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_modular = carriageway_s_modular['length'].sum()

num_of_rows_cw_s_modular_2x96 = carriageway_s_modular_2x96['length'].sum()

num_of_rows_cw_s_modular_total = num_of_rows_cw_s_modular + num_of_rows_cw_s_modular_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["B43"] = num_of_rows_cw_s_modular_total
wb.save(filename)

# *********************** carriageway_S_concrete loc no and wayleave no  *********************************************

carriageway_s_concrete = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Access') ]

carriageway_s_concrete_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Access')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_concrete = carriageway_s_concrete['length'].sum()

num_of_rows_cw_s_concrete_2x96 = carriageway_s_concrete_2x96['length'].sum()

num_of_rows_cw_s_concrete_total = num_of_rows_cw_s_concrete + num_of_rows_cw_s_concrete_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["C43"] = num_of_rows_cw_s_concrete_total
wb.save(filename)

# *********************** carriageway_S_unmade and grassverge loc no and wayleave no  *********************************************

carriageway_s_grassverge = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Access')]

carriageway_s_unmade = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Access')]

carriageway_s_grassverge_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Access')
                         & (df['96mm'] == '2x96')]

carriageway_s_unmade_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Access')
                         & (df['96mm'] == '2x96')]


num_of_rows_cw_s_grassverge = carriageway_s_grassverge['length'].sum()

num_of_rows_cw_s_unmade = carriageway_s_unmade['length'].sum()

num_of_rows_cw_s_grassverge_2x96 = carriageway_s_grassverge_2x96['length'].sum()

num_of_rows_cw_s_unmade_2x96 = carriageway_s_unmade_2x96['length'].sum()

num_of_rows_cw_s_unmade_grassverge_total = num_of_rows_cw_s_grassverge + num_of_rows_cw_s_unmade + num_of_rows_cw_s_grassverge_2x96 + num_of_rows_cw_s_unmade_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["D43"] = num_of_rows_cw_s_unmade_grassverge_total
wb.save(filename)

# *********************** carriageway_S_tarmac loc no and wayleave no  *********************************************

carriageway_s_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac' ) & (df['type'] == 'Access') ]

carriageway_s_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Access')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_s_tarmac = carriageway_s_tarmac['length'].sum()

num_of_rows_cw_s_tarmac_2x96 = carriageway_s_tarmac_2x96['length'].sum()

num_of_rows_cw_s_tarmac_total = num_of_rows_cw_s_tarmac + num_of_rows_cw_s_tarmac_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["E43"] = num_of_rows_cw_s_tarmac_total
wb.save(filename)

# *********************** carriageway_T_modular loc no and wayleave no  *********************************************

carriageway_t_modular = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular' ) & (df['type'] == 'Trunk') ]

carriageway_t_modular_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Modular') & (df['type'] == 'Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_t_modular = carriageway_t_modular['length'].sum()

num_of_rows_cw_t_modular_2x96 = carriageway_t_modular_2x96['length'].sum()

num_of_rows_cw_t_modular_total = num_of_rows_cw_t_modular + num_of_rows_cw_t_modular_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["B44"] = num_of_rows_cw_t_modular_total
wb.save(filename)

# *********************** carriageway_T_concrete loc no and wayleave no  *********************************************

carriageway_t_concrete = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Trunk') ]

carriageway_t_concrete_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Concrete') & (df['type'] == 'Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_t_concrete = carriageway_t_concrete['length'].sum()

num_of_rows_cw_t_concrete_2x96 = carriageway_t_concrete_2x96['length'].sum()

num_of_rows_cw_t_concrete_total = num_of_rows_cw_t_concrete + num_of_rows_cw_t_concrete_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["C44"] = num_of_rows_cw_t_concrete_total
wb.save(filename)

# *********************** carriageway_t_unmade and grassverge loc no and wayleave no  *********************************************

carriageway_t_grassverge = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Trunk')]

carriageway_t_unmade = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Trunk')]

carriageway_t_grassverge_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Grass Verge') & (df['type'] == 'Trunk')
                         & (df['96mm'] == '2x96')]

carriageway_t_unmade_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Unmade') & (df['type'] == 'Trunk')
                         & (df['96mm'] == '2x96')]


num_of_rows_cw_t_grassverge = carriageway_t_grassverge['length'].sum()

num_of_rows_cw_t_unmade = carriageway_t_unmade['length'].sum()

num_of_rows_cw_t_grassverge_2x96 = carriageway_t_grassverge_2x96['length'].sum()

num_of_rows_cw_t_unmade_2x96 = carriageway_t_unmade_2x96['length'].sum()

num_of_rows_cw_t_unmade_grassverge_total = num_of_rows_cw_t_grassverge + num_of_rows_cw_t_unmade + num_of_rows_cw_t_grassverge_2x96 + num_of_rows_cw_t_unmade_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["D44"] = num_of_rows_cw_t_unmade_grassverge_total
wb.save(filename)

# *********************** carriageway_T_tarmac loc no and wayleave no  *********************************************

carriageway_t_tarmac = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac' ) & (df['type'] == 'Trunk') ]

carriageway_t_tarmac_2x96 = df.loc[(df['state'] == 'Planned') & (df['surface'] == 'Carriageway') & (df['loc'] == False)
                         & (df['wayleave'] == False) & (df['material'] == 'Tarmac') & (df['type'] == 'Trunk')
                         & (df['96mm'] == '2x96')]

num_of_rows_cw_t_tarmac = carriageway_t_tarmac['length'].sum()

num_of_rows_cw_t_tarmac_2x96 = carriageway_t_tarmac_2x96['length'].sum()

num_of_rows_cw_t_tarmac_total = num_of_rows_cw_t_tarmac + num_of_rows_cw_t_tarmac_2x96


filename = "Documents/Output/BOQ and BOM.xlsx"

n = 4
wb = load_workbook(filename)
sheets = wb.sheetnames
ws = wb[sheets[n]]
ws_tables = []
ws["E44"] = num_of_rows_cw_t_tarmac_total
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


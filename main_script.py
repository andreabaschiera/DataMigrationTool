from openpyxl import load_workbook
import pandas as pd

from outils.paste_values import paste_values
from outils.access_NR_table import access_NR_table
from outils.copy_values import copy_and_paste

old_wb = load_workbook("Template/VSME-Digital-Template-Sample-1.0.1.xlsx")
new_wb_empty = load_workbook("Template/VSME-Digital-Template-1.1.1.xlsx")

df_old = access_NR_table(old_wb.defined_names)
df_new = access_NR_table(new_wb_empty.defined_names)

df_old_wv = copy_and_paste(old_wb, df_old)
df_old_tomerge = df_old_wv[["name_ranges", "cell_values"]]

df_new_wv = df_new.merge(df_old_tomerge)
df_new_wv = df_new_wv[df_new_wv["name_ranges"].str.startswith("template_") == False]
df_new_wv = df_new_wv[df_new_wv["name_ranges"].str.startswith("enum_") == False]
df_new_wv = df_new_wv[df_new_wv["name_ranges"].str.contains("Table") == False]
df_new_wv = df_new_wv[df_new_wv["name_ranges"].str.contains("Hypercube") == False].reset_index(drop=True) # there's only one "hidden" table (NR: BreakdownOfAnnualMassFlowOfRelevantMaterialsUsedByTheUndertakingHypercube)

paste_values(new_wb_empty, df_new_wv)

new_wb_empty.save("Template/try101.xlsx")

# anti_join = df_old[~df_old['name_ranges'].isin(df_new['name_ranges'])]
# print(anti_join)

# df_new_wv.dropna()

# df_new_wv.to_pickle("df1.pkl")
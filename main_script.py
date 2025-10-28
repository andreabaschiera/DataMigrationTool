from openpyxl import load_workbook
import pandas as pd

from outils.access_NR_table import access_NR_table
from outils.copy_values import copy_and_paste

old_wb = load_workbook("Template/VSME-Digital-Template-Sample-1.0.1.xlsx")
new_wb_empty = load_workbook("Template/VSME-Digital-Template-1.1.1.xlsx")

df_old = access_NR_table(old_wb.defined_names)
df_new = access_NR_table(new_wb_empty.defined_names)

df_old_withvalues = copy_and_paste(old_wb, df_old)
df_old_tomerge = df_old_withvalues[["name_ranges", "cell_values"]]

# df_new_withvalues = df_new.merge(df_old_tomerge)

anti_join = df_old[~df_old['name_ranges'].isin(df_new['name_ranges'])]
print(anti_join)
import pandas as pd

from outils.create_shapes import create_shapes

def access_missingNR_table(df_missingNR, version_cell):
    """ Access sheets, cell references and coordinates of each missing name range"""

    list_of_sheets = []
    list_of_ranges = []

    for i in range(len(df_missingNR)):
        list_of_sheets.append(df_missingNR.loc[:,version_cell].values[i].split("!")[0])
        list_of_ranges.append(df_missingNR.loc[:,version_cell].values[i].split("!")[1])

    for rang in list_of_ranges:
        if rang == "None":
            list_of_ranges[list_of_ranges.index(rang)] = None

    list_of_shapes = create_shapes(list_of_ranges)

    return pd.DataFrame({
        "sheets": list_of_sheets,
        "cell_ranges": list_of_ranges,
        "cell_shapes": list_of_shapes
    })
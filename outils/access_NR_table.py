import pandas as pd

from outils.create_shapes import create_shapes

def access_NR_table(pyxl_NR):
    """ Access names, cell references and coordinates of each name range from the python object containing name ranges"""

    list_NR = list(pyxl_NR.keys())
    list_sheets = []
    list_ranges = []

    for NR in list_NR:
        list_sheets.append(pyxl_NR[NR].attr_text.split("!")[0].replace("'", ""))
        list_ranges.append(pyxl_NR[NR].attr_text.split("!")[1])

    list_shapes = create_shapes(list_ranges)

    return pd.DataFrame({
        "name_ranges": list_NR,
        "sheets": list_sheets,
        "cell_ranges": list_ranges,
        "cell_shapes": list_shapes
    })
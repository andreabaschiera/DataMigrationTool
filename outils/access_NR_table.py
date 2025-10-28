import pandas as pd


def letter_to_number(col):
    """ Convert Excel column letters to numbers (e.g., 'A' -> 1, 'Z' -> 26, 'AA' -> 27) """

    col = col.upper()
    n = 0

    for c in col:
        if not ("A" <= c <= "Z"):
            raise ValueError(f"Invalid column letter: {col}")
        n = n * 26 + (ord(c) - ord("A") + 1)
        
    return n


def access_NR_table(pyxl_NR):
    """ Access names, cell references and coordinates of each name range from the python object containing name ranges"""

    list_NR = list(pyxl_NR.keys())
    list_sheets = []
    list_ranges = []
    list_shapes = []

    for NR in list_NR:
        list_sheets.append(pyxl_NR[NR].attr_text.split("!")[0].replace("'", ""))
        list_ranges.append(pyxl_NR[NR].attr_text.split("!")[1])

    for i in range(len(list_NR)):
        a = list_ranges[i].split("$")[1:]
        for j in range(len(a)):
            if j % 2 != 1:
                a[j] = letter_to_number(a[j])
        if( len(a) == 4 ):
            a[1] = a[1][:-1]
        a = [int(ciao) for ciao in a]
        list_shapes.append(a)

    return pd.DataFrame({
        "name_ranges": list_NR,
        "sheets": list_sheets,
        "cell_ranges": list_ranges,
        "cell_shapes": list_shapes
    })
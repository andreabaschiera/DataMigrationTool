def copy_and_paste(pyxl, df):
    """ Copy and paste values from old workbook to new workbook based on name ranges and cell shapes """

    cell_values = []

    for i in range(len(df)):

        sheet = pyxl[df["sheets"][i]]

        if len(df["cell_shapes"][i]) == 2:
            cell_values.append([sheet[df["cell_ranges"][i]].value])
        else:
            cell_values_temp = []
            for row in sheet.iter_rows(min_row=df["cell_shapes"][i][1],
                                        max_row=df["cell_shapes"][i][3],
                                        min_col=df["cell_shapes"][i][0],
                                        max_col=df["cell_shapes"][i][2]):
                cell_values_temp.append([])
                for col in row:
                    cell_values_temp[-1].append(col.value)
            cell_values.append(cell_values_temp)
    
    df["cell_values"] = cell_values

    return df
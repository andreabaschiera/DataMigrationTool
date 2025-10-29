def copy_and_paste(pyxl, df):
    """ Copy and paste values from old workbook to new workbook based on name ranges and cell shapes """

    cell_values = []

    for i in range(len(df)):

        sheet = pyxl[df["sheets"][i]]

        if len(df["cell_shapes"][i]) == 2:
            if(sheet.cell(row=df["cell_shapes"][i][1], column=df["cell_shapes"][i][0]).data_type != "f"): #checking whether cell value is a formula
                cell_values.append([[sheet[df["cell_ranges"][i]].value]])
            else:
                cell_values.append(None)
        else:
            cell_values_temp = []
            found = False
            for row in sheet.iter_rows(min_row=df["cell_shapes"][i][1],
                                        max_row=df["cell_shapes"][i][3],
                                        min_col=df["cell_shapes"][i][0],
                                        max_col=df["cell_shapes"][i][2]):
                cell_values_temp.append([])
                for col in row:
                    if col.data_type == "f": 
                        cell_values_temp = None
                        found = True
                        break
                    else:
                        cell_values_temp[-1].append(col.value)
                if found:
                    break
            cell_values.append(cell_values_temp)
    
    df["cell_values"] = cell_values

    return df
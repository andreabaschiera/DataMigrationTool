def letter_to_number(col):
    """ Convert Excel column letters to numbers (e.g., 'A' -> 1, 'Z' -> 26, 'AA' -> 27) """

    col = col.upper()
    n = 0

    for c in col:
        if not ("A" <= c <= "Z"):
            raise ValueError(f"Invalid column letter: {col}")
        n = n * 26 + (ord(c) - ord("A") + 1)
        
    return n

def create_shapes(list_of_ranges):
    """ Create cell shapes from cell ranges"""

    list_shapes = []

    for cell_range in list_of_ranges:
        if cell_range is None: # this is for those checkboxes which were created only in newer versions of the template
            list_shapes.append(None)
            continue
        a = cell_range.split("$")[1:]
        for j in range(len(a)):
            if j % 2 != 1:
                a[j] = letter_to_number(a[j])
        if( len(a) == 4 ):
            a[1] = a[1][:-1]
        a = [int(ciao) for ciao in a]
        list_shapes.append(a)

    return list_shapes
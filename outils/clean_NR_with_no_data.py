import pandas as pd

def clean_NR_with_no_data(df):

    list_of_keywords = ["template_", "enum_", "Table", "BreakdownOfEnergyConsumptionAxis", "Hypercube"]

    for kw in list_of_keywords:
        df = df[df["name_ranges"].str.contains(kw) == False]

    return df.reset_index(drop=True)


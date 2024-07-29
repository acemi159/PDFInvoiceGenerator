import glob

import pandas as pd


filepaths = glob.glob("samples/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
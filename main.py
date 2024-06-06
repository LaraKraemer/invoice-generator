import pandas as pd
import glob

filepaths = glob.glob("invoice-generator/invoice/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepaths, sheet_name="Sheet 1")
    print(df)


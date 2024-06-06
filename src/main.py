import pandas as pd
import glob

filepaths = glob.glob("invoices/*.xlsx")
print("Filepaths found:", filepaths)


if not filepaths:
    print("No files found.")
else:
    for filepath in filepaths:
        # Print the filepath being processed
        print("Processing file:", filepath)
        df = pd.read_excel(filepath, sheet_name="Sheet 1")
        print(df)





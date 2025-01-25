import pandas as pd
import glob

filepaths = glob.glob(pathname="invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath)
    print(df)

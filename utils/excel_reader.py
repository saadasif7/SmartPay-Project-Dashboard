import pandas as pd

def load_data():
    df = pd.read_excel("data/projects.xlsx", engine="openpyxl")

    # Remove extra spaces from column names
    df.columns = df.columns.str.strip()

    return df
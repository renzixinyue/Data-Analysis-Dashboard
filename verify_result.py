import pandas as pd
import sys

# Set encoding to utf-8 for stdout
sys.stdout.reconfigure(encoding='utf-8')

def verify():
    filename = 'analysis_result.xlsx'
    print(f"--- {filename} ---")
    try:
        xls = pd.ExcelFile(filename)
        for sheet_name in xls.sheet_names:
            print(f"\nSheet: {sheet_name}")
            df = pd.read_excel(filename, sheet_name=sheet_name)
            print(df.head().to_markdown(index=False, numalign="left", stralign="left"))
            print(f"Shape: {df.shape}")
    except Exception as e:
        print(f"Error reading {filename}: {e}")

if __name__ == "__main__":
    verify()

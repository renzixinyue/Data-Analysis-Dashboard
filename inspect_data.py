import pandas as pd

def print_head(filename):
    print(f"--- {filename} ---")
    try:
        df = pd.read_excel(filename)
        print(df.head().to_markdown(index=False, numalign="left", stralign="left"))
        print("\nColumns:", df.columns.tolist())
    except Exception as e:
        print(f"Error reading {filename}: {e}")
    print("\n")

print_head('diyiciyuekao.xlsx')
print_head('qizhognchengji.xlsx')

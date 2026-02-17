import openpyxl

def print_head_openpyxl(filename):
    print(f"--- {filename} ---")
    try:
        wb = openpyxl.load_workbook(filename, read_only=True, data_only=True)
        sheet = wb.active
        rows = list(sheet.iter_rows(max_row=5, values_only=True))
        for row in rows:
            print(row)
    except Exception as e:
        print(f"Error reading {filename}: {e}")
    print("\n")

print_head_openpyxl('diyiciyuekao.xlsx')
print_head_openpyxl('qizhognchengji.xlsx')

import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx', read_only=True)
print("Sheet names in ITC Model.xlsx:")
for idx, name in enumerate(wb.sheetnames):
    print(f" {idx+1}. {name}")

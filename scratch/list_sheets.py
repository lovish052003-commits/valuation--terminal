import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx', read_only=True)
print("Sheet names in ITC Model.xlsx:")
for i, name in enumerate(wb.sheetnames, 1):
    print(f"  {i:2d}. {name}")
wb.close()

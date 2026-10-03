import openpyxl

wb = openpyxl.load_workbook('scratch/test_dupont_altman_output.xlsx')
dup = wb['Dupont Analysis']
print("Dupont B5:", dup['B5'].value)
print("Dupont B8:", str(dup['B8'].value)[:60])
print("Dupont B37:", str(dup['B37'].value)[:60])
alt = wb["Altman's Z Score"]
print("Altman B5:", alt['B5'].value)
print("Altman B8:", str(alt['B8'].value)[:60])
print("Altman B36:", str(alt['B36'].value)[:60])

import openpyxl

wb = openpyxl.load_workbook('ITC Model.xlsx')

dup = wb['Dupont Analysis']
alt = wb["Altman's Z Score"]

test_about = (
    "Adani Enterprises Limited (AEL) is an Indian multinational publicly listed holding company "
    "and a part of Adani Group. It is headquartered in Ahmedabad and primarily involved in mining and "
    "trading of coal and iron ore. Through its various subsidiaries, it also has business interests in airport "
    "operations, edible oils, road, rail and water infrastructure, data centers, and solar manufacturing."
)

test_updates = [
    "Adani Enterprises secured approval for a major capex deployment in green hydrogen and solar manufacturing ecosystem.",
    "Airport business recorded a 24% YoY growth in passenger traffic across its 7 managed airports in India.",
    "Road and water infrastructure divisions bagged new HAM projects worth over Rs 8,500 crore from NHAI.",
    "Consolidated revenue reached Rs 1,11,431 crore with strong operating EBITDA margin expansion in incubation portfolio.",
    "Management reaffirmed aggressive capital expenditure plans across energy transition, logistics, and digital infrastructure."
]

# Dupont
dup['B5'].value = "52 Week (High - INR - 3,743.00 & low INR - 2,142.00)"
dup['B8'].value = test_about
dup['B37'].value = test_updates[0]
dup['B39'].value = test_updates[1]
dup['B41'].value = test_updates[2]
dup['B43'].value = test_updates[3]
dup['B45'].value = test_updates[4]

# Altman
alt['B5'].value = "52 Week (High - INR - 3,743.00 & low INR - 2,142.00)"
alt['B8'].value = test_about
alt['B36'].value = test_updates[0]
alt['B38'].value = test_updates[1]
alt['B40'].value = test_updates[2]
alt['B42'].value = test_updates[3]
alt['B44'].value = test_updates[4]

wb.save('scratch/test_dupont_altman_output.xlsx')
wb.close()
print("Saved scratch/test_dupont_altman_output.xlsx successfully!")

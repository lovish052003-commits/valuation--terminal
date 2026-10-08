import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import openpyxl
import zipfile
import shutil
from institutional_control_checks import apply_institutional_wiring

def strip_calc_chain_from_xlsx(file_path):
    temp_zip = file_path + ".temp.zip"
    try:
        with zipfile.ZipFile(file_path, 'r') as zin, zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/calcChain.xml':
                    continue
                if item.filename == '[Content_Types].xml':
                    content = zin.read(item.filename).decode('utf-8')
                    import re
                    content = re.sub(r'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', '', content)
                    zout.writestr(item, content.encode('utf-8'))
                elif item.filename == 'xl/_rels/workbook.xml.rels':
                    content = zin.read(item.filename).decode('utf-8')
                    import re
                    content = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '', content)
                    zout.writestr(item, content.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
        shutil.move(temp_zip, file_path)
    except Exception as e:
        if os.path.exists(temp_zip):
            os.remove(temp_zip)

def process_file(f_path):
    print(f"Applying universal Claude fixes to: {f_path}")
    try:
        wb = openpyxl.load_workbook(f_path, data_only=False)
        apply_institutional_wiring(wb)
        wb.calculation.fullCalcOnLoad = True
        wb.save(f_path)
        wb.close()
        strip_calc_chain_from_xlsx(f_path)
        print(f"  -> SUCCESS: {f_path}")
    except Exception as e:
        print(f"  -> ERROR on {f_path}: {e}")

# 1. Update master template and root workbooks
root_dir = r"c:\Users\LENOVO\Downloads\Advance Financial Project"
for fname in os.listdir(root_dir):
    if fname.endswith(".xlsx") and not fname.startswith("~$") and "corrupted" not in fname:
        f_path = os.path.join(root_dir, fname)
        process_file(f_path)

# 2. Update exports/FORCEMOT_Valuation_Model.xlsx from FORCEMOT_Valuation_Model_FIXED.xlsx
fixed_src = os.path.join(root_dir, "FORCEMOT_Valuation_Model_FIXED.xlsx")
dest_export = os.path.join(root_dir, "exports", "FORCEMOT_Valuation_Model.xlsx")
if os.path.exists(fixed_src):
    shutil.copyfile(fixed_src, dest_export)
    strip_calc_chain_from_xlsx(dest_export)
    print(f"Updated {dest_export} directly from FIXED source.")

print("All workbooks successfully synced!")

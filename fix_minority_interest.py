"""
fix_minority_interest.py

Universal Python script using openpyxl that automatically patches the missing
Minority Interest deduction in the Enterprise-to-Equity Value bridge for any
valuation model Excel file.

Exact Logic Implemented:
1. Find Minority Interest Value: Loads workbook with data_only=False, searches
   Column 2 (B) of 'Raw FS' for 'controlling' or 'minority', scans left-to-right
   to extract the last (most recent) numerical value.
2. Store in Data Sheet: Sets cell A73 to 'Minority Interest' and cell K73 to the
   numerical value found.
3. Locate DCF Bridge: Dynamically finds the row labeled exactly 'Equity Value' in
   Column B of 'DCF' sheet (eq_row).
4. Insert Deduction Row: Inserts a new row directly above 'Equity Value' via
   ws.insert_rows(eq_row).
5. Populate New Row: Sets Column B to 'Less: Minority Interest' and Column D to
   formula "='Data Sheet'!K73".
6. Update Equity Value Formula: Shifts to eq_row + 1. Appends -D[NewlyInsertedRowNumber]
   (e.g., -D39) to the existing formula so Minority Interest is deducted.
   Also dynamically updates downstream Equity Value per Share to reference the
   shifted cells.
7. Save & Clean: Saves as [Original_Filename]_Final.xlsx, and strips stale calcChain.xml
   to guarantee zero Excel repair/recovery warnings on load.
"""

import os
import sys
import re
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def strip_calc_chain_from_xlsx(xlsx_path):
    """
    Strips xl/calcChain.xml from the workbook package so Excel opens cleanly
    without any 'We found a problem with some content...' recovery warnings.
    Excel rebuilds the calculation chain natively upon opening.
    """
    if not xlsx_path or not os.path.exists(xlsx_path):
        return
    temp_zip = xlsx_path + ".stripcalc.tmp"
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zin:
            if 'xl/calcChain.xml' not in zin.namelist():
                return
            with zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'xl/calcChain.xml':
                        continue
                    data = zin.read(item.filename)
                    if item.filename == '[Content_Types].xml':
                        data = re.sub(rb'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b'', data)
                    elif item.filename == 'xl/_rels/workbook.xml.rels':
                        data = re.sub(rb'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', b'', data)
                    zout.writestr(item, data)
        os.replace(temp_zip, xlsx_path)
    except Exception as e:
        if os.path.exists(temp_zip):
            try:
                os.remove(temp_zip)
            except Exception:
                pass
        print(f"[Warning] strip_calc_chain notice: {e}")


def fix_minority_interest(excel_filename: str) -> str:
    """
    Patches the Enterprise-to-Equity bridge in the given Excel workbook by inserting
    Minority Interest from Raw FS into Data Sheet and DCF.

    Args:
        excel_filename (str): Path to the input Excel workbook.

    Returns:
        str: Path to the saved [Original_Filename]_Final.xlsx workbook.
    """
    if not os.path.exists(excel_filename):
        raise FileNotFoundError(f"Input file not found: {excel_filename}")

    print(f"\n==================================================================")
    print(f"PATCHING MINORITY INTEREST FOR: {excel_filename}")
    print(f"==================================================================")

    # Load workbook preserving formulas (data_only=False)
    wb = openpyxl.load_workbook(excel_filename, data_only=False)

    # -------------------------------------------------------------------------
    # STEP 1: Find the Minority Interest Value from 'Raw FS'
    # -------------------------------------------------------------------------
    if 'Raw FS' not in wb.sheetnames:
        raise ValueError(f"'Raw FS' sheet not found in {excel_filename}")

    ws_raw = wb['Raw FS']
    minority_val = 0.0
    found_row = None
    found_label = ""

    for r in range(1, ws_raw.max_row + 1):
        cell_val = ws_raw.cell(row=r, column=2).value
        if cell_val and any(term in str(cell_val).lower() for term in ['controlling', 'minority']):
            found_row = r
            found_label = str(cell_val).strip()
            break

    if found_row:
        # Scan row from left to right (col 3 onwards) to find last numerical value
        for c in range(3, ws_raw.max_column + 1):
            v = ws_raw.cell(row=found_row, column=c).value
            if v is not None:
                if isinstance(v, (int, float)):
                    minority_val = float(v)
                elif isinstance(v, str):
                    try:
                        minority_val = float(v.replace(',', '').strip())
                    except ValueError:
                        # Reached text label of adjacent table (e.g. 'Net Profit +')
                        break
            else:
                # If we have already captured values and hit an empty gap cell, stop
                if minority_val != 0.0:
                    break
        print(f"[Step 1] Found '{found_label}' on Raw FS Row {found_row} -> Latest Value: {minority_val}")
    else:
        print(f"[Step 1] No 'controlling' or 'minority' label found in Raw FS Column B. Defaulting to 0.0")

    # -------------------------------------------------------------------------
    # STEP 2: Store in Data Sheet (Cell A73 = 'Minority Interest', Cell K73 = value)
    # -------------------------------------------------------------------------
    if 'Data Sheet' not in wb.sheetnames:
        raise ValueError(f"'Data Sheet' not found in {excel_filename}")

    ws_data = wb['Data Sheet']
    ws_data['A73'] = "Minority Interest"
    ws_data['K73'] = minority_val
    print(f"[Step 2] Stored in Data Sheet: Cell A73 = '{ws_data['A73'].value}', Cell K73 = {ws_data['K73'].value}")

    # -------------------------------------------------------------------------
    # STEP 3: Locate DCF Bridge ('Equity Value' in Column B)
    # -------------------------------------------------------------------------
    if 'DCF' not in wb.sheetnames:
        raise ValueError(f"'DCF' sheet not found in {excel_filename}")

    ws_dcf = wb['DCF']
    eq_row = None

    for r in range(1, ws_dcf.max_row + 1):
        cell_val = ws_dcf.cell(row=r, column=2).value
        if cell_val and str(cell_val).strip().lower() == 'equity value':
            eq_row = r
            break

    if not eq_row:
        raise ValueError("Could not find row labeled 'Equity Value' in Column B of DCF sheet")

    print(f"[Step 3] Located DCF 'Equity Value' dynamically at Row {eq_row}")

    # -------------------------------------------------------------------------
    # STEP 4: Insert New Row Directly Above 'Equity Value'
    # -------------------------------------------------------------------------
    ws_dcf.insert_rows(eq_row)
    print(f"[Step 4] Inserted new row at Row {eq_row} (shifted previous rows down by 1)")

    # -------------------------------------------------------------------------
    # STEP 5: Populate the New Row (Row eq_row)
    # -------------------------------------------------------------------------
    ws_dcf.cell(row=eq_row, column=2, value="Less: Minority Interest")
    ws_dcf.cell(row=eq_row, column=4, value="='Data Sheet'!K73")

    # Match styling with row directly above (e.g. Less: Debt)
    above_row = eq_row - 1
    cell_b_above = ws_dcf.cell(row=above_row, column=2)
    cell_d_above = ws_dcf.cell(row=above_row, column=4)

    if cell_b_above.font:
        ws_dcf.cell(row=eq_row, column=2).font = Font(
            name=cell_b_above.font.name,
            size=cell_b_above.font.size,
            bold=cell_b_above.font.bold,
            italic=cell_b_above.font.italic,
            color=cell_b_above.font.color
        )
    if cell_b_above.alignment:
        ws_dcf.cell(row=eq_row, column=2).alignment = Alignment(
            horizontal=cell_b_above.alignment.horizontal,
            vertical=cell_b_above.alignment.vertical
        )
    if cell_d_above.number_format:
        ws_dcf.cell(row=eq_row, column=4).number_format = cell_d_above.number_format

    print(f"[Step 5] Populated Row {eq_row}: Col B = '{ws_dcf.cell(row=eq_row, column=2).value}', Col D = '{ws_dcf.cell(row=eq_row, column=4).value}'")

    # -------------------------------------------------------------------------
    # STEP 6: Update Equity Value Formula (Now Shifted to eq_row + 1)
    # -------------------------------------------------------------------------
    shifted_eq_row = eq_row + 1
    existing_formula = str(ws_dcf.cell(row=shifted_eq_row, column=4).value or "")

    # Append -D{eq_row} so Minority Interest in cell D{eq_row} is properly subtracted
    subtraction_term = f"-D{eq_row}"
    if subtraction_term not in existing_formula:
        if existing_formula.startswith('='):
            new_eq_formula = f"{existing_formula}{subtraction_term}"
        else:
            new_eq_formula = f"={existing_formula}{subtraction_term}"
        ws_dcf.cell(row=shifted_eq_row, column=4, value=new_eq_formula)
    else:
        new_eq_formula = existing_formula

    print(f"[Step 6] Updated Equity Value Formula at Row {shifted_eq_row}:")
    print(f"         Previous: {existing_formula}")
    print(f"         Updated : {new_eq_formula}")

    # Synchronize downstream 'Equity Value per Share' formula
    # Find row labeled 'Equity Value per Share' below shifted_eq_row
    for r in range(shifted_eq_row + 1, shifted_eq_row + 6):
        lbl = ws_dcf.cell(row=r, column=2).value
        if lbl and 'equity value per share' in str(lbl).strip().lower():
            shares_row = shifted_eq_row + 1
            ws_dcf.cell(row=r, column=4, value=f"=D{shifted_eq_row}/D{shares_row}")
            print(f"         Synchronized 'Equity Value per Share' at Row {r}: =D{shifted_eq_row}/D{shares_row}")
            break

    # Synchronize downstream 'Discount / Premium' formula
    for r in range(shifted_eq_row + 1, shifted_eq_row + 10):
        lbl = str(ws_dcf.cell(row=r, column=2).value or '').strip().lower()
        if 'discount' in lbl or 'premium' in lbl:
            ws_dcf.cell(row=r, column=4, value=f"=D{r-1}/D{r-3}")
            print(f"         Synchronized 'Discount / Premium' at Row {r}: =D{r-1}/D{r-3}")
            break

    # -------------------------------------------------------------------------
    # STEP 7: Save as [Original_Filename]_Final.xlsx
    # -------------------------------------------------------------------------
    base_name, ext = os.path.splitext(excel_filename)
    if base_name.endswith('_Final'):
        output_filename = excel_filename
    else:
        output_filename = f"{base_name}_Final{ext}"

    wb.calculation.fullCalcOnLoad = True
    wb.save(output_filename)
    wb.close()

    # Clean calculation chain to eliminate Excel corruption / repair warnings
    strip_calc_chain_from_xlsx(output_filename)

    print(f"[Step 7] Successfully saved patched workbook to:")
    print(f"         -> {output_filename}")
    print(f"==================================================================\n")
    return output_filename


if __name__ == '__main__':
    # Accept filename from CLI argument or default to latest model
    if len(sys.argv) > 1:
        target_file = sys.argv[1]
    else:
        default_candidate = os.path.join('exports', 'SUNPHARMA_Valuation_Model.xlsx')
        if os.path.exists(default_candidate):
            target_file = default_candidate
        else:
            target_file = 'ITC Model.xlsx'

    output_path = fix_minority_interest(target_file)
    print(f"Result: {output_path}")

import openpyxl
from copy import copy

def patch_master_template():
    wb_tata = openpyxl.load_workbook('Tata Steel Final Model.xlsx')
    ws_tata = wb_tata['Comp_Valuation']

    wb_master = openpyxl.load_workbook('master_model_template.xlsx')
    ws_master = wb_master['Comp_Valuation']

    # 1. Update headers
    ws_master['I10'].value = "EV/Revenue"
    ws_master['J10'].value = "EV/EBITDA"
    ws_master['O10'].value = "EV/Revenue"
    ws_master['P10'].value = "EV/EBITDA"
    ws_master['Q10'].value = "P/E"
    ws_master['N10'].value = None

    # Copy header styling from O10 to I10, J10
    ref_hdr = ws_tata['O10']
    for col in ['I', 'J', 'O', 'P', 'Q']:
        cell = ws_master[f'{col}10']
        if ref_hdr.has_style:
            cell.font = copy(ref_hdr.font)
            cell.border = copy(ref_hdr.border)
            cell.fill = copy(ref_hdr.fill)
            cell.number_format = copy(ref_hdr.number_format)
            cell.alignment = copy(ref_hdr.alignment)

    # 2. Update peer rows (12 to 21)
    for idx in range(10):
        r = 12 + idx
        r_raw = 57 + idx

        ws_master[f'B{r}'].value = f"='Raw FS'!L{r_raw}"
        ws_master[f'D{r}'].value = f"='Raw FS'!M{r_raw}"
        ws_master[f'E{r}'].value = f"='Raw FS'!N{r_raw}"
        ws_master[f'F{r}'].value = f"=D{r}*E{r}"
        ws_master[f'G{r}'].value = f"='Raw FS'!AS{r_raw}"
        ws_master[f'H{r}'].value = f"='Raw FS'!AT{r_raw}"
        ws_master[f'I{r}'].value = f'=IF(OR(ISBLANK(K{r}), K{r}<=0), "N/A", IFERROR($H{r}/K{r}, "N/A"))'
        ws_master[f'J{r}'].value = f'=IF(OR(ISBLANK(L{r}), L{r}<=0), "N/A", IFERROR($H{r}/L{r}, "N/A"))'
        ws_master[f'K{r}'].value = f"='Raw FS'!AU{r_raw}"
        ws_master[f'L{r}'].value = f"='Raw FS'!AV{r_raw}"
        ws_master[f'M{r}'].value = f"='Raw FS'!AW{r_raw}"
        ws_master[f'N{r}'].value = None
        ws_master[f'O{r}'].value = f'=IF(OR(ISBLANK(K{r}), K{r}<=0), "N/A", IFERROR($H{r}/K{r}, "N/A"))'
        ws_master[f'P{r}'].value = f'=IF(OR(ISBLANK(L{r}), L{r}<=0), "N/A", IFERROR($H{r}/L{r}, "N/A"))'
        ws_master[f'Q{r}'].value = f'=IF(OR(ISBLANK(M{r}), M{r}<=0), "N/A", IFERROR(F{r}/M{r}, "N/A"))'

        # Apply styles from Tata Steel row 12
        for col in ['B','C','D','E','F','G','H','I','J','K','L','M','O','P','Q']:
            t_cell = ws_tata[f'{col}12']
            m_cell = ws_master[f'{col}{r}']
            if t_cell.has_style:
                m_cell.font = copy(t_cell.font)
                m_cell.border = copy(t_cell.border)
                m_cell.fill = copy(t_cell.fill)
                m_cell.number_format = copy(t_cell.number_format)
                m_cell.alignment = copy(t_cell.alignment)

    # 3. Update benchmark statistics (Rows 23 to 28)
    p_rng_i = "I12:I21"
    p_rng_j = "J12:J21"
    p_rng_o = "O12:O21"
    p_rng_p = "P12:P21"
    p_rng_q = "Q12:Q21"

    benchmarks = [
        (23, "High", "MAX"),
        (24, "25th Percentile", "_xlfn.QUARTILE.INC"),
        (25, "Median", "MEDIAN"),
        (26, "Average", "AVERAGE"),
        (27, "75th Percentile", "_xlfn.QUARTILE.INC"),
        (28, "Low", "MIN")
    ]

    for row_idx, label, fn in benchmarks:
        ws_master[f'B{row_idx}'].value = label
        if fn == "_xlfn.QUARTILE.INC":
            arg_suffix = ",1)" if label == "25th Percentile" else ",3)"
            ws_master[f'I{row_idx}'].value = f"={fn}({p_rng_i}{arg_suffix}"
            ws_master[f'J{row_idx}'].value = f"={fn}({p_rng_j}{arg_suffix}"
            ws_master[f'O{row_idx}'].value = f"={fn}({p_rng_o}{arg_suffix}"
            ws_master[f'P{row_idx}'].value = f"={fn}({p_rng_p}{arg_suffix}"
            ws_master[f'Q{row_idx}'].value = f"={fn}({p_rng_q}{arg_suffix}"
        else:
            ws_master[f'I{row_idx}'].value = f"={fn}({p_rng_i})"
            ws_master[f'J{row_idx}'].value = f"={fn}({p_rng_j})"
            ws_master[f'O{row_idx}'].value = f"={fn}({p_rng_o})"
            ws_master[f'P{row_idx}'].value = f"={fn}({p_rng_p})"
            ws_master[f'Q{row_idx}'].value = f"={fn}({p_rng_q})"

        for col in ['B', 'I', 'J', 'O', 'P', 'Q']:
            t_cell = ws_tata[f'{col}{row_idx}']
            m_cell = ws_master[f'{col}{row_idx}']
            if t_cell.has_style:
                m_cell.font = copy(t_cell.font)
                m_cell.border = copy(t_cell.border)
                m_cell.fill = copy(t_cell.fill)
                m_cell.number_format = copy(t_cell.number_format)
                m_cell.alignment = copy(t_cell.alignment)

    # 4. Target Valuation Section (Rows 30 to 39)
    ws_master['B30'].value = "='Raw FS'!L56 & \" Comparable Valuation\""
    ws_master['I30'].value = "EV/Revenue"
    ws_master['J30'].value = "EV/EBITDA"
    ws_master['O30'].value = "EV/Revenue"
    ws_master['P30'].value = "EV/EBITDA"
    ws_master['Q30'].value = "P/E"

    # Row 32: Implied EV
    ws_master['B32'].value = "Implied Enterprise Value"
    ws_master['I32'].value = "='Raw FS'!AU56*I25"
    ws_master['J32'].value = "='Raw FS'!AV56*J25"
    ws_master['O32'].value = "='Raw FS'!AU56*O25"
    ws_master['P32'].value = "='Raw FS'!AV56*P25"
    ws_master['Q32'].value = "='Raw FS'!AW56*Q25"

    # Row 33: Net Debt
    ws_master['B33'].value = "Net Debt Value"
    ws_master['I33'].value = "='Raw FS'!AS56"
    ws_master['J33'].value = "='Raw FS'!AS56"
    ws_master['O33'].value = "='Raw FS'!AS56"
    ws_master['P33'].value = "='Raw FS'!AS56"
    ws_master['Q33'].value = 0

    # Row 34: Implied Market Value
    ws_master['B34'].value = "Implied Market Value"
    ws_master['I34'].value = "=I32-I33"
    ws_master['J34'].value = "=J32-J33"
    ws_master['O34'].value = "=O32-O33"
    ws_master['P34'].value = "=P32-P33"
    ws_master['Q34'].value = "=Q32"

    # Row 35: Share Outstanding
    ws_master['B35'].value = "Share Outstanding"
    ws_master['I35'].value = "='Raw FS'!N56"
    ws_master['J35'].value = "='Raw FS'!N56"
    ws_master['O35'].value = "='Raw FS'!N56"
    ws_master['P35'].value = "='Raw FS'!N56"
    ws_master['Q35'].value = "='Raw FS'!N56"

    # Row 37: Implied Value per Share
    ws_master['B37'].value = "Implied Value per Share"
    ws_master['I37'].value = "=I34/I35"
    ws_master['J37'].value = "=J34/J35"
    ws_master['O37'].value = "=O34/O35"
    ws_master['P37'].value = "=P34/P35"
    ws_master['Q37'].value = "=Q34/Q35"

    # Row 38: Source
    ws_master['B38'].value = "Source : Screener.in"

    # Row 39: Verdict
    ws_master['B39'].value = "Verdict"
    ws_master['I39'].value = '=IF(I37="N/A","N/A",IF(I37>\'Raw FS\'!M56,"Undervalued","Overvalued"))'
    ws_master['J39'].value = '=IF(J37="N/A","N/A",IF(J37>\'Raw FS\'!M56,"Undervalued","Overvalued"))'
    ws_master['O39'].value = '=IF(O37="N/A","N/A",IF(O37>\'Raw FS\'!M56,"Undervalued","Overvalued"))'
    ws_master['P39'].value = '=IF(P37="N/A","N/A",IF(P37>\'Raw FS\'!M56,"Undervalued","Overvalued"))'
    ws_master['Q39'].value = '=IF(Q37="N/A","N/A",IF(Q37>\'Raw FS\'!M56,"Undervalued","Overvalued"))'

    # Copy styles for rows 30 to 39
    for r in [30, 32, 33, 34, 35, 37, 38, 39]:
        for col in ['B', 'I', 'J', 'O', 'P', 'Q']:
            t_cell = ws_tata[f'{col}{r}']
            m_cell = ws_master[f'{col}{r}']
            if t_cell.has_style:
                m_cell.font = copy(t_cell.font)
                m_cell.border = copy(t_cell.border)
                m_cell.fill = copy(t_cell.fill)
                m_cell.number_format = copy(t_cell.number_format)
                m_cell.alignment = copy(t_cell.alignment)

    # Column widths to avoid ### overflow
    for col_letter in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'O', 'P', 'Q']:
        ws_master.column_dimensions[col_letter].width = 16.5

    wb_master.save('master_model_template.xlsx')
    print("Successfully updated master_model_template.xlsx Comp_Valuation sheet from Tata Steel Final Model standard.")

if __name__ == '__main__':
    patch_master_template()

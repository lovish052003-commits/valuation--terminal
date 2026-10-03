import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

from run_master_institutional_test import test_company

def run_all_sectors():
    sectors = [
        ('OIL', False, 'Energy / Oil & Gas'),
        ('SUNPHARMA', False, 'Pharma / Healthcare'),
        ('TATASTEEL', False, 'Manufacturing / Metals'),
        ('TCS', False, 'IT Services / Tech')
    ]
    
    summary = {}
    for ticker, is_bank, sector_desc in sectors:
        print(f"\n{'='*70}")
        print(f"TESTING SECTOR ADAPTABILITY: {ticker} ({sector_desc})")
        print(f"{'='*70}")
        try:
            ok = test_company(ticker, is_bank=is_bank)
            summary[f"{ticker} ({sector_desc})"] = "PASSED" if ok else "FAILED"
        except Exception as e:
            print(f"❌ Error testing {ticker}: {e}")
            import traceback
            traceback.print_exc()
            summary[f"{ticker} ({sector_desc})"] = f"ERROR: {e}"
            
    print("\n" + "="*70)
    print("MULTI-SECTOR INSTITUTIONAL REGRESSION SUMMARY:")
    for k, v in summary.items():
        print(f"   {k}: {v}")
    print("="*70)

if __name__ == '__main__':
    run_all_sectors()

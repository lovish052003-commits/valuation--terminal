"""Adds the separate Alt Valuation Excel model to app.py (additive; .bak made; aborts on mismatch)."""
import sys, shutil
p = sys.argv[1] if len(sys.argv) > 1 else 'app.py'
s = open(p, encoding='utf-8').read()
if 'alt_valuation_export' in s:
    print('already patched'); sys.exit(0)

def rep(old, new):
    global s
    assert s.count(old) == 1, f'expected 1 match, found {s.count(old)} for: {old[:60]!r}'
    s = s.replace(old, new)

# 1. import
rep("EXPORT_STATUS = {}\n", "EXPORT_STATUS = {}\nALT_STATUS = {}\ntry:\n    import alt_valuation_export\nexcept Exception as _e_alt:\n    alt_valuation_export = None\n    print(f'[Alt Model] module unavailable: {_e_alt}')\n")

# 2. generate after the DCF engine runs
rep("        # 3. AI Report Generation\n",
"""        # 2b. Alternate valuation workbook for companies where the FCFF DCF does not apply
        alt_info = None
        if alt_valuation_export is not None:
            try:
                _iv = val_result.get('intrinsic_value_per_share')
                _ok, _why = alt_valuation_export.is_dcf_applicable(screener_data, _iv)
                _txt = ' '.join(str(val_result.get(k, '')) for k in ('verdict', 'dcf_status', 'notice')).upper()
                if 'NOT APPLICABLE' in _txt or 'NOTICE' in _txt:
                    _ok, _why = False, _why or 'DCF flagged not applicable by engine'
                if not _ok:
                    _r = alt_valuation_export.export_alt_model(screener_data, excel_exporter.EXPORT_DIR)
                    alt_info = {k: v for k, v in _r.items() if k != 'path'}
                    alt_info['reason'] = _why
                    alt_info['download_url'] = f"/api/download-alt/{_r['filename']}"
                    ALT_STATUS[ticker] = dict(_r, reason=_why)
                    val_result['alt_model'] = alt_info
                else:
                    ALT_STATUS.pop(ticker, None)
            except Exception as _e_alt:
                print(f'[Alt Model] generation failed for {ticker}: {_e_alt}')

        # 3. AI Report Generation
""")

# 3. return in the JSON
rep("            'excel_filename': excel_filename,\n            'raw_tables': raw_tables,",
    "            'excel_filename': excel_filename,\n            'alt_model': alt_info,\n            'raw_tables': raw_tables,")

# 4. routes
rep("@app.route('/api/download-excel/<filename>')\n",
"""@app.route('/api/alt-model/<ticker>')
def alt_model_info(ticker):
    t = ticker.upper()
    info = ALT_STATUS.get(t)
    if not info:
        return jsonify({'available': False})
    out = {k: v for k, v in info.items() if k != 'path'}
    out['available'] = True
    out['download_url'] = f"/api/download-alt/{info['filename']}"
    return jsonify(sanitize_for_json(out))

@app.route('/api/download-alt/<filename>')
def download_alt(filename):
    clean = os.path.basename(filename)
    path = os.path.join(excel_exporter.EXPORT_DIR, clean)
    if not clean.endswith('_Alt_Valuation_Model.xlsx') or not os.path.exists(path):
        return jsonify({'error': 'Alternate valuation model not found'}), 404
    return send_file(path, as_attachment=True, download_name=clean)

@app.route('/api/download-excel/<filename>')
""")
shutil.copy(p, p + '.bak_alt')
open(p, 'w', encoding='utf-8').write(s)
print('OK patched', p)

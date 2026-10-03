import numpy as np
from screener_client import fetch_company_data
from excel_exporter import build_wacc_peer_companies

data = fetch_company_data('ADANIPOWER')
comps = build_wacc_peer_companies(data)

eff_tax_rate = 0.30
rf = 0.068
erp = 0.065
pre_tax_kd = 0.0565
post_tax_kd = pre_tax_kd * (1.0 - eff_tax_rate)

unlev_betas = []
peer_wds = []
for c in comps[:5]:
    c_debt = float(c.get('debt') or 0.0)
    c_mcap = float(c.get('mcap') or 1.0)
    c_beta = float(c.get('beta') or 1.0)
    de_ratio = c_debt / c_mcap if c_mcap > 0 else 0.0
    u_b = c_beta / (1.0 + (1.0 - eff_tax_rate) * de_ratio)
    unlev_betas.append(u_b)
    if c_debt + c_mcap > 0:
        peer_wds.append(c_debt / (c_debt + c_mcap))

unlevered_beta = float(np.median(unlev_betas))
target_wd = float(np.median(peer_wds))
target_we = 1.0 - target_wd
target_de = target_wd / target_we
levered_beta = unlevered_beta * (1.0 + (1.0 - eff_tax_rate) * target_de)
ke = rf + levered_beta * erp
wacc = (target_we * ke) + (target_wd * post_tax_kd)

print(f"unlevered_beta: {unlevered_beta:.4f}")
print(f"target_wd: {target_wd:.4f}, target_we: {target_we:.4f}, target_de: {target_de:.4f}")
print(f"levered_beta: {levered_beta:.4f}")
print(f"ke: {ke*100:.2f}%")
print(f"post_tax_kd: {post_tax_kd*100:.2f}%")
print(f"WACC: {wacc*100:.2f}%")

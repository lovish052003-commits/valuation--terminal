import os

c_name = "Adani Enterprises Ltd"
ticker = "ADANIENT"
STOP_WORDS = {
    'ltd', 'limited', 'india', 'corp', 'corporation', 'industries', 'industry',
    'enterprises', 'enterprise', 'holdings', 'holding', 'services', 'service',
    'solutions', 'technologies', 'technology', 'finance', 'financial', 'motors',
    'motor', 'products', 'product', 'global', 'group', 'company', 'co'
}
search_dirs = [
    os.path.expanduser(r'~\Downloads'),
    r"C:\Users\LENOVO\Downloads",
    os.path.dirname(os.path.abspath('.')),
    os.path.abspath('.'),
    os.path.abspath('exports')
]
search_tokens = set()
if ticker:
    search_tokens.add(ticker.lower())
    for sfx in ['ind', 'ltd', 'corp']:
        if ticker.lower().endswith(sfx) and len(ticker) > len(sfx) + 2:
            search_tokens.add(ticker.lower()[:-len(sfx)])
for w in c_name.replace('(', '').replace(')', '').split():
    wl = w.lower()
    if len(wl) >= 3 and wl not in STOP_WORDS:
        search_tokens.add(wl)

print("Search tokens:", search_tokens)
for sdir in search_dirs:
    if not os.path.exists(sdir):
        continue
    for fname in os.listdir(sdir):
        if not fname.endswith('.xlsx') or fname.startswith('~$') or fname.startswith('.'):
            continue
        fl = fname.lower()
        for tok in search_tokens:
            if tok in fl:
                print(f"Found in {sdir}: {fname}")
                break

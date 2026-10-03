import re

def is_company_match(p_name, c_name, ticker=''):
    if not p_name or not c_name:
        return False
    def clean_tokens(s):
        s = re.sub(r'[^\w\s]', '', s.lower())
        stopwords = {'ltd', 'limited', 'inds', 'industries', 'ind', 'india', 'co', 'corp', 'corporation'}
        return [t for t in s.split() if t not in stopwords]
    
    tp = clean_tokens(p_name)
    tc = clean_tokens(c_name)
    if not tp or not tc:
        return False
    
    # 1. Exact token match
    if tp == tc:
        return True
    
    # 2. Ticker match
    clean_p = re.sub(r'[^a-z0-9]', '', p_name.lower())
    clean_c = re.sub(r'[^a-z0-9]', '', c_name.lower())
    clean_t = re.sub(r'[^a-z0-9]', '', ticker.lower()) if ticker else ''
    
    if clean_p == clean_c:
        return True
    if clean_t and (clean_t == clean_p or clean_t in clean_p):
        return True
        
    # 3. Subsets (require at least 2 matching tokens, or if single token it must be exact)
    set_p = set(tp)
    set_c = set(tc)
    if set_p == set_c:
        return True
    if len(tc) >= 2 and set_c.issubset(set_p):
        return True
    if len(tp) >= 2 and set_p.issubset(set_c):
        return True
    if len(tc) == 1 and len(tp) == 1 and tc[0] == tp[0]:
        return True
        
    return False

test_peers = [
    'JSW Steel',
    'Tata Steel',
    'Jindal Steel',
    'S A I L',
    'Jindal Stain.',
    'Sarda Energy',
    'Vedanta Iron & Steel',
    'Tata Consumer Products',
    'Tata Motors',
    'Tata Power'
]

for p in test_peers:
    res = is_company_match(p, "Tata Steel Ltd", "TATASTEEL")
    print(f"{p:25} -> {res}")

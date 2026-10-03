import re
import openpyxl

def is_company_match(p_name, c_name, ticker):
    if not p_name or not c_name:
        return False
    def clean(s):
        s = re.sub(r'[^\w\s]', '', s.lower())
        for word in ['ltd', 'limited', 'inds', 'industries', 'ind', 'india', 'co', 'corp', 'corporation']:
            s = re.sub(r'\b' + word + r'\b', '', s)
        return re.sub(r'\s+', ' ', s).strip()
    
    cp = clean(p_name)
    cc = clean(c_name)
    if cp == cc or cp in cc or cc in cp:
        return True
    words_p = set(cp.split())
    words_c = set(cc.split())
    if words_p and words_c and (words_p.issubset(words_c) or words_c.issubset(words_p)):
        return True
    if ticker and ticker.lower() in p_name.lower().replace(' ', ''):
        return True
    return False

def build_peer_range(col, target_row):
    """
    Builds the range string excluding target_row from rows 12:21.
    """
    if target_row <= 12:
        return f"{col}13:{col}21"
    elif target_row >= 21:
        return f"{col}12:{col}20"
    elif target_row == 13:
        return f"({col}12, {col}14:{col}21)"
    elif target_row == 20:
        return f"({col}12:{col}19, {col}21)"
    else:
        return f"({col}12:{col}{target_row-1}, {col}{target_row+1}:{col}21)"

# Test cases
print("Matching Tests:")
print("  Tata Steel vs Tata Steel Ltd:", is_company_match("Tata Steel", "Tata Steel Ltd", "TATASTEEL"))
print("  JSW Steel vs Tata Steel Ltd:", is_company_match("JSW Steel", "Tata Steel Ltd", "TATASTEEL"))
print("  Hind. Unilever vs Hindustan Unilever Ltd:", is_company_match("Hind. Unilever", "Hindustan Unilever Ltd", "HINDUNILVR"))
print("  ITC vs ITC Ltd:", is_company_match("ITC", "ITC Ltd", "ITC"))
print("  Britannia Inds. vs Britannia Industries Ltd:", is_company_match("Britannia Inds.", "Britannia Industries Ltd", "BRITANNIA"))

print("\nRange Generation Tests:")
for tr in [12, 13, 14, 17, 20, 21]:
    r_str = build_peer_range("P", tr)
    print(f"  Target Row {tr} -> Peer Range: {r_str}")

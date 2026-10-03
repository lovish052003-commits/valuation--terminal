import os, json, base64, sqlite3, shutil
import win32crypt
from Crypto.Cipher import AES

def get_chrome_cookies(domain='screener.in'):
    local_state_path = os.path.expanduser(r'~\AppData\Local\Google\Chrome\User Data\Local State')
    cookies_path = os.path.expanduser(r'~\AppData\Local\Google\Chrome\User Data\Default\Network\Cookies')
    
    if not os.path.exists(local_state_path) or not os.path.exists(cookies_path):
        print("Chrome cookie files not found.")
        return {}
        
    try:
        with open(local_state_path, 'r', encoding='utf-8') as f:
            local_state = json.load(f)
        encrypted_key = base64.b64decode(local_state['os_crypt']['encrypted_key'])[5:]
        key = win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)[1]
    except Exception as e:
        print("Could not decrypt master key:", e)
        return {}

    tmp_cookies = 'scratch_cookies.db'
    shutil.copy2(cookies_path, tmp_cookies)
    conn = sqlite3.connect(tmp_cookies)
    c = conn.cursor()
    c.execute("SELECT name, encrypted_value FROM cookies WHERE host_key LIKE ?", (f'%{domain}%',))
    
    cookies = {}
    for name, enc_val in c.fetchall():
        try:
            nonce = enc_val[3:15]
            ciphertext = enc_val[15:-16]
            tag = enc_val[-16:]
            cipher = AES.new(key, AES.MODE_GCM, nonce=nonce)
            decrypted = cipher.decrypt_and_verify(ciphertext, tag).decode('utf-8')
            cookies[name] = decrypted
        except Exception as e:
            pass
            
    conn.close()
    if os.path.exists(tmp_cookies):
        os.remove(tmp_cookies)
    return cookies

c = get_chrome_cookies('screener.in')
print("Extracted Screener cookies:", list(c.keys()))
if 'sessionid' in c:
    print("Found sessionid!")

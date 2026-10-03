import re

with open('static/css/style.css', 'r', encoding='utf-8') as f:
    text = f.read()

# remove :root and [data-theme="dark"]
pos_root = text.find(':root')
end_root = text.find('}', pos_root)

pos_dark = text.find('[data-theme="dark"]')
end_dark = text.find('}', pos_dark)

body_css = text[:pos_root] + text[end_root+1:pos_dark] + text[end_dark+1:]

hexes = set(re.findall(r'#[0-9a-fA-F]{3,6}\b', body_css))
print("Hex colors in main stylesheet:")
for h in sorted(list(hexes)):
    # print occurrences
    lines = [idx+1 for idx, l in enumerate(text.splitlines()) if h in l and idx+1 > end_dark]
    print(f"  {h}: lines {lines[:5]}")

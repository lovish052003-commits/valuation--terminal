with open('static/css/style.css', 'r', encoding='utf-8') as f:
    lines = f.readlines()

print(f"Total lines: {len(lines)}")
hardcoded = []
for i, line in enumerate(lines, 1):
    if i <= 46:  # :root definitions
        continue
    # look for hardcoded hex or rgba backgrounds
    if any(k in line for k in ['background:', 'background-color:', 'color: #', 'background: #', 'background: rgba(9,', 'background: rgba(15,', 'background: rgba(21,']):
        hardcoded.append((i, line.strip()))

print(f"Found {len(hardcoded)} potential hardcoded color lines.")
for line_no, text in hardcoded[:40]:
    print(f"L{line_no:4d}: {text}")

import re

with open('static/css/style.css', 'r', encoding='utf-8') as f:
    css = f.read()

# 1. Update :root and add [data-theme="dark"]
new_root_and_dark = """
:root {
  /* Default: Clean, High-Contrast White / Light Theme */
  --bg-primary: #F8FAFC;
  --bg-secondary: #FFFFFF;
  --bg-card: #FFFFFF;
  --bg-card-hover: #F1F5F9;
  --bg-elevated: #F1F5F9;
  --bg-glass: rgba(255, 255, 255, 0.92);
  --bg-header: rgba(255, 255, 255, 0.92);
  --bg-input: #FFFFFF;
  --bg-input-focus: #FFFFFF;
  --bg-dropdown: #FFFFFF;
  --bg-item-hover: #F1F5F9;
  --bg-subtle: #F8FAFC;
  --bg-toolbar: #F8FAFC;
  --bg-api-badge: #F1F5F9;
  --bg-verdict: linear-gradient(135deg, #FFFFFF 0%, #F1F5F9 100%);

  --border-subtle: #E2E8F0;
  --border-medium: #CBD5E1;
  --border-active: #0284C7;
  --border-glow: 0 0 15px rgba(2, 132, 199, 0.15);

  --text-primary: #0F172A;
  --text-secondary: #334155;
  --text-muted: #64748B;
  --text-inverse: #FFFFFF;

  --title-gradient: linear-gradient(135deg, #0F172A 30%, #2563EB 100%);
  --table-header-bg: #F8FAFC;
  --table-row-hover: #F1F5F9;
  --table-border: #E2E8F0;

  --slider-track: #E2E8F0;
  --modal-bg: #FFFFFF;
  --modal-backdrop: rgba(15, 23, 42, 0.5);
  --toast-bg: #0F172A;
  --toast-text: #FFFFFF;

  --ff-track: #E2E8F0;
  --formula-bg: #F8FAFC;
  --pm-bg: #F8FAFC;
  --badge-bg: rgba(2, 132, 199, 0.08);
  --chip-bg: #F1F5F9;
  --chip-hover-bg: #E2E8F0;

  --accent-emerald: #059669;
  --accent-emerald-glow: rgba(5, 150, 105, 0.15);
  --accent-amber: #D97706;
  --accent-amber-glow: rgba(217, 119, 6, 0.15);
  --accent-rose: #E11D48;
  --accent-rose-glow: rgba(225, 29, 72, 0.15);
  --accent-cyan: #0284C7;
  --accent-blue: #2563EB;
  --accent-purple: #7C3AED;

  --font-sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  --font-mono: 'JetBrains Mono', monospace;

  --radius-sm: 6px;
  --radius-md: 10px;
  --radius-lg: 16px;
  --radius-full: 9999px;

  --shadow-sm: 0 1px 2px 0 rgba(15, 23, 42, 0.05);
  --shadow-md: 0 4px 12px -1px rgba(15, 23, 42, 0.06), 0 2px 4px -2px rgba(15, 23, 42, 0.04);
  --shadow-lg: 0 10px 25px -3px rgba(15, 23, 42, 0.08), 0 4px 6px -4px rgba(15, 23, 42, 0.03);
  --shadow-glow: 0 0 20px rgba(2, 132, 199, 0.2);

  --bg-body-gradient: radial-gradient(circle at 15% 15%, rgba(37, 99, 235, 0.04) 0%, transparent 45%),
                      radial-gradient(circle at 85% 75%, rgba(14, 165, 233, 0.04) 0%, transparent 45%);
}

[data-theme="dark"] {
  /* Antigravity Sleek Cosmic Dark Mode */
  --bg-primary: #090D16;
  --bg-secondary: #0F172A;
  --bg-card: rgba(21, 29, 46, 0.85);
  --bg-card-hover: rgba(30, 41, 59, 0.95);
  --bg-elevated: #1E293B;
  --bg-glass: rgba(30, 41, 59, 0.6);
  --bg-header: rgba(9, 13, 22, 0.85);
  --bg-input: #0F172A;
  --bg-input-focus: #111B31;
  --bg-dropdown: #0F172A;
  --bg-item-hover: #1E293B;
  --bg-subtle: rgba(15, 23, 42, 0.6);
  --bg-toolbar: #111A2E;
  --bg-api-badge: rgba(15, 23, 42, 0.6);
  --bg-verdict: linear-gradient(135deg, rgba(30, 41, 59, 0.9), rgba(15, 23, 42, 0.9));

  --border-subtle: rgba(255, 255, 255, 0.08);
  --border-medium: rgba(255, 255, 255, 0.15);
  --border-active: rgba(56, 189, 248, 0.4);
  --border-glow: 0 0 15px rgba(56, 189, 248, 0.15);

  --text-primary: #F8FAFC;
  --text-secondary: #94A3B8;
  --text-muted: #64748B;
  --text-inverse: #090D16;

  --title-gradient: linear-gradient(135deg, #FFFFFF 30%, #94A3B8 100%);
  --table-header-bg: #111927;
  --table-row-hover: rgba(255, 255, 255, 0.02);
  --table-border: rgba(255, 255, 255, 0.04);

  --slider-track: #1E293B;
  --modal-bg: #0F172A;
  --modal-backdrop: rgba(0, 0, 0, 0.75);
  --toast-bg: #1E293B;
  --toast-text: #FFFFFF;

  --ff-track: #111927;
  --formula-bg: #0F172A;
  --pm-bg: rgba(15, 23, 42, 0.6);
  --badge-bg: rgba(56, 189, 248, 0.15);
  --chip-bg: rgba(30, 41, 59, 0.5);
  --chip-hover-bg: #1E293B;

  --accent-emerald: #10B981;
  --accent-emerald-glow: rgba(16, 185, 129, 0.2);
  --accent-amber: #F59E0B;
  --accent-amber-glow: rgba(245, 158, 11, 0.2);
  --accent-rose: #F43F5E;
  --accent-rose-glow: rgba(244, 63, 94, 0.2);
  --accent-cyan: #38BDF8;
  --accent-blue: #3B82F6;
  --accent-purple: #8B5CF6;

  --shadow-sm: 0 1px 3px rgba(0, 0, 0, 0.3);
  --shadow-md: 0 4px 12px rgba(0, 0, 0, 0.4);
  --shadow-lg: 0 10px 30px -5px rgba(0, 0, 0, 0.6);
  --shadow-glow: 0 0 25px rgba(56, 189, 248, 0.25);

  --bg-body-gradient: radial-gradient(circle at 15% 15%, rgba(56, 189, 248, 0.05) 0%, transparent 40%),
                      radial-gradient(circle at 85% 75%, rgba(139, 92, 246, 0.05) 0%, transparent 40%);
}
"""

# Replace existing :root block
old_root_pattern = r':root\s*\{[^}]*\}'
css = re.sub(old_root_pattern, new_root_and_dark.strip(), css, count=1)

# Body background
old_body_bg = r'background-image:\s*radial-gradient\(circle at 15% 15%[^\)]*\)[^;]*;'
css = re.sub(old_body_bg, 'background-image: var(--bg-body-gradient);', css, count=1)

# Header background
css = css.replace('background: rgba(9, 13, 22, 0.85);', 'background: var(--bg-header);')

# Api badge background
css = css.replace('background: rgba(15, 23, 42, 0.6);', 'background: var(--bg-api-badge);')

# Command title gradient
css = css.replace('background: linear-gradient(135deg, #FFFFFF 30%, #94A3B8 100%);', 'background: var(--title-gradient);')

# Search input focus background
css = css.replace('background: #111B31;', 'background: var(--bg-input-focus);')

# Autocomplete dropdown background & item
css = css.replace('background: #0F172A;', 'background: var(--bg-dropdown);')
css = css.replace('border-bottom: 1px solid rgba(255, 255, 255, 0.04);', 'border-bottom: 1px solid var(--border-subtle);')
css = css.replace('background: #1E293B;', 'background: var(--bg-item-hover);')

# Chips
css = css.replace('background: rgba(30, 41, 59, 0.5);', 'background: var(--chip-bg);')

# Verdict card
css = css.replace('background: linear-gradient(135deg, rgba(30, 41, 59, 0.9), rgba(15, 23, 42, 0.9));', 'background: var(--bg-verdict);')

# Params toolbar
css = css.replace('background: #111A2E;', 'background: var(--bg-toolbar);')
css = css.replace('border: 1px solid rgba(56, 189, 248, 0.2);', 'border: 1px solid var(--border-subtle);')

# Range slider track
# Note: #1E293B was replaced in autocomplete item, let's verify range slider
css = re.sub(r'(\.range-slider\s*\{[^}]*background:\s*)#1E293B;', r'\1var(--slider-track);', css)

# Data table th and td
css = css.replace('background: #111927;', 'background: var(--table-header-bg);')
css = css.replace('background: rgba(255, 255, 255, 0.02);', 'background: var(--table-row-hover);')

# Report viewer th
css = css.replace('background: #151D2E;', 'background: var(--table-header-bg);')

# Modal card
# Replace background: var(--bg-dropdown); inside modal-card to var(--modal-bg) if needed, or modal
css = re.sub(r'(\.modal-card\s*\{[^}]*background:\s*)var\(--bg-dropdown\);', r'\1var(--modal-bg);', css)

# Form input
css = re.sub(r'(\.form-select,\s*\.form-input\s*\{[^}]*background:\s*)var\(--bg-item-hover\);', r'\1var(--bg-elevated);', css)

# Toast
css = re.sub(r'(\.toast\s*\{[^}]*background:\s*)var\(--bg-item-hover\);', r'\1var(--toast-bg);', css)
css = re.sub(r'(\.toast\s*\{[^}]*color:\s*)white;', r'\1var(--toast-text);', css)

# Football field track
css = re.sub(r'(\.ff-bar-wrap\s*\{[^}]*background:\s*)var\(--table-header-bg\);', r'\1var(--ff-track);', css)

# Add Antigravity Theme Switcher Toggle styles & smooth transition
theme_switcher_styles = """
/* Antigravity Theme Switcher Toggle */
.theme-toggle-btn {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  background: var(--bg-elevated);
  border: 1px solid var(--border-subtle);
  color: var(--text-secondary);
  padding: 0.35rem 0.75rem;
  border-radius: var(--radius-full);
  font-size: 0.82rem;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.25s ease;
  user-select: none;
}

.theme-toggle-btn:hover {
  color: var(--text-primary);
  border-color: var(--accent-cyan);
  background: var(--bg-card-hover);
  box-shadow: var(--shadow-glow);
}

.theme-icon {
  font-size: 0.95rem;
  line-height: 1;
  display: inline-block;
  transition: transform 0.3s ease;
}

.theme-toggle-btn:hover .theme-icon {
  transform: rotate(15deg) scale(1.1);
}

.theme-text {
  font-size: 0.8rem;
  font-weight: 600;
  letter-spacing: -0.01em;
}

.theme-toggle-pill {
  width: 34px;
  height: 18px;
  background: var(--border-medium);
  border-radius: var(--radius-full);
  position: relative;
  transition: background 0.25s ease;
}

[data-theme="dark"] .theme-toggle-pill {
  background: var(--accent-cyan);
}

.theme-toggle-dot {
  position: absolute;
  top: 2px;
  left: 2px;
  width: 14px;
  height: 14px;
  background: #FFFFFF;
  border-radius: 50%;
  transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.2);
}

[data-theme="dark"] .theme-toggle-dot {
  transform: translateX(16px);
  background: #090D16;
}

/* Global Theme Transition */
body, header, .glass-panel, .search-input-wrapper, .params-toolbar, .scorecard, .pillar-card, .data-table th, .data-table td, .modal-card, .btn-action, .chip, .formula-flow, .pm-item, .autocomplete-dropdown, .form-select, .form-input, .theme-toggle-btn, .theme-toggle-pill {
  transition: background-color 0.25s ease, border-color 0.25s ease, color 0.25s ease, box-shadow 0.25s ease;
}
"""

css += "\n" + theme_switcher_styles

with open('static/css/style.css', 'w', encoding='utf-8') as f:
    f.write(css)

print("Successfully updated static/css/style.css!")

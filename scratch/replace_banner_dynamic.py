import os
import glob
import zipfile
import shutil
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

# 1. Create the new banner image
canvas = Image.new('RGBA', (357, 360), (0, 0, 0, 0))
draw = ImageDraw.Draw(canvas)

# Visible area is roughly Y=64 to Y=288 due to Excel crop settings.
# Draw the teal line on the left
draw.rectangle([80, 80, 85, 270], fill=(0, 204, 153, 255))

# Draw the text
try:
    font = ImageFont.truetype("arialbd.ttf", 36)
except Exception:
    font = ImageFont.load_default()

text = "DCF\nVALUATION\nMODEL"
draw.multiline_text((105, 85), text, fill=(255, 255, 255, 255), font=font, spacing=15)

buf = BytesIO()
canvas.save(buf, format='PNG')
img_bytes = buf.getvalue()

# Save locally to view if needed
with open('scratch/new_banner_fixed.png', 'wb') as f:
    f.write(img_bytes)

# 2. Replace the banner image in all Excel templates
files = glob.glob('*.xlsx')
for f in files:
    if f.startswith('~'):
        continue
    temp_path = f + '.tmp'
    updated = False
    try:
        with zipfile.ZipFile(f, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename.startswith('xl/media/') and item.filename.endswith('.png'):
                        # Check size
                        img_data = zin.read(item.filename)
                        try:
                            im = Image.open(BytesIO(img_data))
                            if im.size == (357, 360):
                                zout.writestr(item, img_bytes)
                                updated = True
                                continue
                        except Exception:
                            pass
                    
                    zout.writestr(item, zin.read(item.filename))
        
        if updated:
            shutil.move(temp_path, f)
            print(f"Updated banner in {f}")
        else:
            os.remove(temp_path)
            print(f"No 357x360 banner found in {f}")
    except Exception as e:
        print(f"Error processing {f}: {e}")

print("Banner replacement complete.")

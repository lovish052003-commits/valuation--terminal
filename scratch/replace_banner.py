import os
import glob
import zipfile
import shutil
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

# 1. Create the new banner image
# Size: 357 x 360, transparent background
canvas = Image.new('RGBA', (357, 360), (0, 0, 0, 0))
draw = ImageDraw.Draw(canvas)

# Draw the teal line on the left
# Color teal: #00b050 roughly from the screenshot, wait, it looks like cyan/teal.
# Let's use (0, 176, 80) which is standard excel green, or (0, 204, 153)
draw.rectangle([5, 10, 12, 280], fill=(0, 204, 153, 255))

# Draw the text: DCF \n VALUATION \n MODEL
try:
    font = ImageFont.truetype("arialbd.ttf", 46) # Bold font
except Exception:
    font = ImageFont.load_default()

text = "DCF\nVALUATION\nMODEL"
draw.multiline_text((25, 20), text, fill=(255, 255, 255, 255), font=font, spacing=15)

buf = BytesIO()
canvas.save(buf, format='PNG')
img_bytes = buf.getvalue()

# Save a copy to see it
with open('scratch/new_banner.png', 'wb') as f:
    f.write(img_bytes)

# 2. Replace image2.png in all Excel templates
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
                    if item.filename == 'xl/media/image2.png':
                        zout.writestr(item, img_bytes)
                        updated = True
                    else:
                        zout.writestr(item, zin.read(item.filename))
        
        if updated:
            shutil.move(temp_path, f)
            print(f"Updated banner in {f}")
        else:
            os.remove(temp_path)
            print(f"No image2.png found in {f}")
    except Exception as e:
        print(f"Error processing {f}: {e}")

print("Banner replacement complete.")

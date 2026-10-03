import os
import glob
import win32com.client
import zipfile
import shutil
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

# 1. Create the new banner image (Sensible padding)
canvas = Image.new('RGBA', (357, 360), (0, 0, 0, 0))
draw = ImageDraw.Draw(canvas)
draw.rectangle([5, 10, 10, 180], fill=(0, 204, 153, 255))
try:
    font = ImageFont.truetype("arialbd.ttf", 36)
except Exception:
    font = ImageFont.load_default()
text = "DCF\nVALUATION\nMODEL"
draw.multiline_text((25, 15), text, fill=(255, 255, 255, 255), font=font, spacing=10)
buf = BytesIO()
canvas.save(buf, format='PNG')
img_bytes = buf.getvalue()

# 2. Inject image
files = glob.glob('*.xlsx')
for f in files:
    if f.startswith('~'): continue
    temp_path = f + '.tmp'
    updated = False
    try:
        with zipfile.ZipFile(f, 'r') as zin, zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename.startswith('xl/media/') and item.filename.endswith('.png'):
                    try:
                        im = Image.open(BytesIO(zin.read(item.filename)))
                        if im.size == (357, 360):
                            zout.writestr(item, img_bytes)
                            updated = True
                            continue
                    except: pass
                zout.writestr(item, zin.read(item.filename))
        if updated:
            shutil.move(temp_path, f)
        else:
            os.remove(temp_path)
    except Exception as e:
        print(f"Zip error on {f}: {e}")

# 3. Uncrop in Excel COM
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
for f in files:
    if f.startswith('~'): continue
    try:
        wb = excel.Workbooks.Open(os.path.abspath(f), UpdateLinks=0)
        for ws in wb.Sheets:
            for shp in ws.Shapes:
                if shp.Type == 13: # msoPicture
                    # If it's roughly the banner size or top left
                    if shp.Top < 50 and shp.Left < 50:
                        shp.PictureFormat.CropTop = 0
                        shp.PictureFormat.CropLeft = 0
                        shp.PictureFormat.CropBottom = 0
                        shp.PictureFormat.CropRight = 0
                        # Also fix aspect ratio and scale to fit original shape size
        wb.Save()
        wb.Close(SaveChanges=True)
        print(f"Uncropped {f}")
    except Exception as e:
        print(f"COM error on {f}: {e}")

excel.Quit()
print("Done")

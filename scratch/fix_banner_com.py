import os
import glob
import win32com.client
from PIL import Image, ImageDraw, ImageFont

# 1. Create a properly proportioned new banner image
canvas = Image.new('RGBA', (580, 280), (0, 0, 0, 0))
draw = ImageDraw.Draw(canvas)
draw.rectangle([15, 30, 25, 250], fill=(0, 204, 153, 255))

try:
    font = ImageFont.truetype("arialbd.ttf", 64)
except Exception:
    font = ImageFont.load_default()

text = "DCF\nVALUATION\nMODEL"
draw.multiline_text((55, 30), text, fill=(255, 255, 255, 255), font=font, spacing=15)

banner_path = os.path.abspath('scratch/new_banner_clean.png')
canvas.save(banner_path, format='PNG')

# 2. Iterate through Excel templates and replace the picture
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

files = glob.glob('*.xlsx')
for f in files:
    if f.startswith('~'): continue
    try:
        path = os.path.abspath(f)
        wb = excel.Workbooks.Open(path, UpdateLinks=0)
        modified = False
        
        for ws in wb.Sheets:
            # We must iterate backwards when deleting items from a collection
            shapes_to_replace = []
            for i in range(1, ws.Shapes.Count + 1):
                shp = ws.Shapes(i)
                # Check if it's a picture and near the top left
                if shp.Type == 13 and shp.Top < 50 and shp.Left < 50:
                    shapes_to_replace.append({
                        'Left': shp.Left,
                        'Top': shp.Top,
                        'Width': shp.Width,
                        'Height': shp.Height,
                        'Name': shp.Name
                    })
            
            for sinfo in shapes_to_replace:
                try:
                    ws.Shapes(sinfo['Name']).Delete()
                    # Add picture: filename, link_to_file, save_with_document, left, top, width, height
                    ws.Shapes.AddPicture(banner_path, False, True, sinfo['Left'], sinfo['Top'], sinfo['Width'], sinfo['Height'])
                    modified = True
                except Exception as e:
                    print(f"Error replacing shape on {ws.Name}: {e}")

        if modified:
            print(f"Replaced banners in {f}")
        
        wb.Save()
        wb.Close(SaveChanges=True)
    except Exception as e:
        print(f"Error processing {f}: {e}")

excel.Quit()
print("Banner COM replacement complete.")

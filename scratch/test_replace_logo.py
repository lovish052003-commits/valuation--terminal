import zipfile
import os
import shutil
from io import BytesIO
from PIL import Image

def test_replace_xlsx_images(xlsx_path, new_img_path):
    temp_path = xlsx_path + ".temp"
    
    # Load and prepare new image
    with Image.open(new_img_path) as im:
        im = im.convert('RGBA')
        
        # Prepare for image3 (330x344 target canvas)
        canvas3 = Image.new('RGBA', (330, 344), (0, 0, 0, 0))
        # resize im keeping aspect ratio
        im_copy = im.copy()
        im_copy.thumbnail((320, 330), Image.Resampling.LANCZOS)
        # paste centered
        x = (330 - im_copy.width) // 2
        y = (344 - im_copy.height) // 2
        canvas3.paste(im_copy, (x, y), im_copy)
        buf3 = BytesIO()
        canvas3.save(buf3, format='PNG')
        img3_bytes = buf3.getvalue()
        
        # Prepare for image2 (357x360 target canvas)
        canvas2 = Image.new('RGBA', (357, 360), (0, 0, 0, 0))
        im_copy2 = im.copy()
        im_copy2.thumbnail((345, 350), Image.Resampling.LANCZOS)
        x2 = (357 - im_copy2.width) // 2
        y2 = (360 - im_copy2.height) // 2
        canvas2.paste(im_copy2, (x2, y2), im_copy2)
        buf2 = BytesIO()
        canvas2.save(buf2, format='PNG')
        img2_bytes = buf2.getvalue()

    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/media/image3.png':
                    zout.writestr(item, img3_bytes)
                    print("Replaced xl/media/image3.png!")
                elif item.filename == 'xl/media/image2.png':
                    zout.writestr(item, img2_bytes)
                    print("Replaced xl/media/image2.png!")
                else:
                    zout.writestr(item, zin.read(item.filename))
                    
    shutil.move(temp_path, xlsx_path)
    print("Successfully updated workbook package:", xlsx_path)

# Test with Tata Steel logo
import requests

r = requests.get('https://icon.horse/icon/tatasteel.com')
with open('scratch/tata_logo.png', 'wb') as f:
    f.write(r.content)

xlsx_test = os.path.abspath('exports/TATASTEEL_Valuation_Model.xlsx')
test_replace_xlsx_images(xlsx_test, 'scratch/tata_logo.png')

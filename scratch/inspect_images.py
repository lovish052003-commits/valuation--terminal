import os
try:
    from PIL import Image
    for name in ['image1.jpeg', 'image2.png', 'image3.png']:
        path = os.path.join('scratch', 'extracted_media', 'xl', 'media', name)
        if os.path.exists(path):
            with Image.open(path) as img:
                print(f'{name}: size={img.size}, mode={img.mode}, format={img.format}')
        else:
            print(f'{path} does not exist')
except Exception as e:
    print('Error:', e)

import os
import pandas as pd
import requests
from PIL import Image
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
import json

excel_file = 'list_urls.xlsx'
image_folder = 'downloaded_images'
output_file = 'list_urls_with_images.xlsx'

os.makedirs(image_folder, exist_ok=True)

df = pd.read_excel(excel_file)

images = []
metadata = []

for i, url in enumerate(df['URL']):
    try:
        if pd.isna(url):
            images.append('')
            metadata.append('')
            continue

        response = requests.get(url, timeout=10)
        image = Image.open(BytesIO(response.content)).convert("RGB")

        filename = f'image_{i}.png'
        filepath = os.path.join(image_folder, filename)
        image.save(filepath)

        file_size_kb = round(os.path.getsize(filepath) / 1024, 2)
        width, height = image.size

        images.append(filepath)

        info = {
            "filename": filename,
            "size_kb": file_size_kb,
            "dimensions": f"{width}x{height}"
        }
        metadata.append(json.dumps(info))

    except Exception as e:
        print(f"Error at row {i}: {e}")
        images.append('')
        metadata.append('')

df['Image_path'] = images
df['Metadata'] = metadata
df.to_excel(output_file, index=False)

wb = load_workbook(output_file)
ws = wb.active

ws.column_dimensions['A'].width = 60  
ws.column_dimensions['B'].width = 30  
ws.column_dimensions['C'].width = 70  

for i, path in enumerate(images):
    row = i + 2
    if path and os.path.exists(path):
        img = ExcelImage(path)
        img.width = 80
        img.height = 80
        ws.add_image(img, f'B{row}')
        ws.row_dimensions[row].height = 60

wb.save(output_file)
print("All Done!")

import subprocess
subprocess.run(['start', output_file], shell=True)
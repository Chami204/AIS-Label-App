import os
import barcode
from barcode.writer import ImageWriter

def generate_barcodes(df, save_dir):
    for _, row in df.iterrows():
        part_number = str(row.get('part_number', '')).strip()
        no = str(row.get('no')).strip()
        if part_number:
            filename = f"{no}.{part_number}"
            code128 = barcode.get('code128', part_number, writer=ImageWriter())
            code128.save(os.path.join(save_dir, filename))

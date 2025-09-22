import math
import pandas as pd
import qrcode
import os
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

# Nama file Excel dan nama sheet
# excel_file = "list-link-2.xlsx"
excel_file = "test.xlsx"
sheet_name = "Sheet1"
colom_name = "Links"
kanwil_qr = "Kanwil"
cabang_qr = "Cabang"
# area_qr = "Area"
output_folder = "qrcodes"
nomor = "NO"
kode_cb = "Kode Cabang"

# Baca file Excel
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Pastikan ada kolom bernama "Link"
if colom_name not in df.columns:
    raise ValueError("Pastikan file Excel memiliki kolom 'Link'")

# Buat folder untuk menyimpan QR Code
os.makedirs(output_folder, exist_ok=True)

# Loop melalui setiap baris di Excel dan buat QR Code
for index , row in df.iterrows():
    link = row[colom_name]
    cabang_name = row[cabang_qr]
    kanwil_name = row[kanwil_qr]
    no = row[nomor]
    kode_cabang = row[kode_cb]

    if pd.isna(cabang_name):
        cabang_name = "Null"

    if pd.isna(kanwil_name):
        kanwil_name = "NULL"
    
    # Lewati baris kosong
    if pd.isna(link):
        print(f"skip" + kanwil_name + cabang_name)
        continue
    
    # Generate QR Code
    qr = qrcode.make(link)

    # Convert QR Code ke format PIL Image
    qr = qr.convert("RGB")

    # Tambahkan watermark
    draw = ImageDraw.Draw(qr)
    text = str(kode_cabang) + " - " + cabang_name  # Watermark

    # Mulai dari ukuran font besar, turunkan jika tidak muat
    max_width = qr.size[0] - 20  # Maks lebar teks (beri padding 10px kiri-kanan)
    font_size = qr.size[0] // 10

    while font_size > 10:
        #windows
        # try:
        #     font = ImageFont.truetype("arial.ttf", font_size)
        # except:
        #     font = ImageFont.load_default()
        #mac
        font_path = "/System/Library/Fonts/Supplemental/Arial.ttf"
        if not Path(font_path).exists():
            font = ImageFont.load_default()
        else:
            font = ImageFont.truetype(font_path, font_size)

        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]

        if text_width <= max_width:
            break  # Ukuran cukup, keluar loop
        font_size -= 1  # Perkecil font

    # Hitung ulang posisi setelah dapat font yang pas
    text_height = text_bbox[3] - text_bbox[1]
    text_x = (qr.size[0] - text_width) / 2
    text_y = qr.size[1] - text_height - 10

    draw.text((text_x, text_y), text, font=font, fill=(255, 0, 0))
    
    # Simpan QR Code dengan nama berdasarkan indeks atau isi link
    qr_filename = os.path.join(output_folder, str(no)+" " +kanwil_name+"-"+cabang_name+f".png")
    qr.save(qr_filename)
    
    print(f"QR Code disimpan sebagai {qr_filename}")

print("✅ Semua QR Code berhasil dibuat!")
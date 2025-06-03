import math
import pandas as pd
import qrcode
import os
from PIL import Image, ImageDraw, ImageFont

# Nama file Excel dan nama sheet
excel_file = "list-link-2.xlsx"
# excel_file = "test.xlsx"
sheet_name = "Sheet1"
colom_name = "Links"
kanwil_qr = "Kanwil"
cabang_qr = "Cabang"
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
    font_size = 20
    font_size = qr.size[0] // 15  # Font size proporsional dengan ukuran QR

    try:
        font = ImageFont.truetype("arial.ttf", font_size)  # Windows
    except:
        font = ImageFont.load_default()  # Linux/Mac tanpa arial.ttf

    text = str(kode_cabang) + " - " + cabang_name  # Watermark
    text_width, text_height = draw.font(text, font=font)

    # # Posisi teks di kiri bawah
    text_x = 10  # Beri jarak 10px dari kiri
    text_y = 580  # Beri jarak 10px dari bawah

    # Ambil ukuran teks untuk penempatan yang akurat
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]

    # # Posisi teks di kiri bawah, dengan padding
    padding_x = 10
    padding_y = 10
    text_x = padding_x  # 10px dari kiri
    text_y = qr.size[1] - text_height - padding_y  # 10px dari bawah

    # Posisi teks di tengah bawah, dengan padding 10px
    text_x = (qr.size[0] - text_width) / 2  # Tengah
    text_y = qr.size[1] - text_height - 10  # 10px dari bawah

    draw.text((text_x, text_y), text, font=font, fill=(0, 255, 0))  # Warna hijau (RGB)
    draw.text((text_x, text_y), text, font=font, fill=(255, 0, 0))  # Warna merah (RGB)
    
    # Simpan QR Code dengan nama berdasarkan indeks atau isi link
    qr_filename = os.path.join(output_folder, str(no)+" " +kanwil_name+"-"+cabang_name+f".png")
    qr.save(qr_filename)
    
    print(f"QR Code disimpan sebagai {qr_filename}")

print("✅ Semua QR Code berhasil dibuat!")
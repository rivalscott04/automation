import os
import re
from PyPDF2 import PdfReader
import shutil

# Path ke folder yang berisi file PDF
source_folder_path = r"C:\Users\Rival\Documents\P3K_T1\SPP_P3K"
renamed_folder_path =  r"C:\Users\Rival\Documents\P3K_T1\SPP_P3K\renamed"

# Membuat folder baru untuk file yang sudah di-rename jika belum ada
if not os.path.exists(renamed_folder_path):
    os.makedirs(renamed_folder_path)

# Fungsi untuk mengambil NIP dan Nama dari file PDF
def extract_nip_and_name(pdf_path):
    reader = PdfReader(pdf_path)
    # Ambil halaman pertama (atau sesuaikan jika perlu)
    page = reader.pages[0]
    text = page.extract_text()

    # Temukan posisi kalimat "Dengan ini menyatakan sesungguhnya bahwa:"
    match_start = re.search(r"Dengan ini menyatakan sesungguhnya bahwa:", text)

    if match_start:
        # Ambil teks setelah kalimat tersebut
        text_after = text[match_start.end():]

        # Ekstrak NIP setelah "NIP :"
        nip_pattern = r"NIP\s*[:\-]?\s*(\d{8,})"  # NIP biasanya berupa angka panjang (8 digit atau lebih)
        nip_match = re.search(nip_pattern, text_after)

        # Ekstrak Nama setelah "Nama :"
        name_pattern = r"Nama\s*[:\-]?\s*(.*)"  # Ambil seluruh teks setelah "Nama :"
        name_match = re.search(name_pattern, text_after)

        if nip_match and name_match:
            nip = nip_match.group(1)
            name = name_match.group(1).strip()
            return nip, name
    return None, None

# Fungsi untuk mengecek apakah file dengan nama yang sama sudah ada
def get_unique_filename(path):
    if not os.path.exists(path):
        return path
    else:
        base, extension = os.path.splitext(path)
        counter = 1
        while os.path.exists(f"{base}({counter}){extension}"):
            counter += 1
        return f"{base}({counter}){extension}"

# Fungsi untuk merename dan memindahkan file PDF
def rename_and_move_pdfs(source_folder_path, renamed_folder_path):
    for filename in os.listdir(source_folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(source_folder_path, filename)
            nip, name = extract_nip_and_name(pdf_path)

            if nip and name:
                # Format nama baru
                new_name = f"SPMT_{nip}_{name}.pdf"
                new_pdf_path = os.path.join(renamed_folder_path, new_name)

                # Cek apakah nama file sudah ada, jika ada, tambahkan (1), (2), dst.
                unique_pdf_path = get_unique_filename(new_pdf_path)

                # Pindahkan dan rename file ke folder baru
                shutil.copy(pdf_path, unique_pdf_path)
                print(f"Renamed and moved: {filename} -> {os.path.basename(unique_pdf_path)}")
            else:
                print(f"Failed to extract NIP and Name from {filename}")

# Jalankan fungsi untuk merename dan memindahkan PDF
rename_and_move_pdfs(source_folder_path, renamed_folder_path)
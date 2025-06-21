import os
import re
from PyPDF2 import PdfReader
from collections import Counter

# Path ke dua folder yang berisi file PDF
folder_a_path = r"C:\Users\Rival\Downloads\SPMT_CPNS_Signed\renamed"
folder_b_path = r"C:\Users\Rival\Documents\SPMT_CPNS"

# Fungsi untuk mengambil Nama dari file PDF
def extract_name(pdf_path):
    reader = PdfReader(pdf_path)
    # Ambil halaman pertama (atau sesuaikan jika perlu)
    page = reader.pages[0]
    text = page.extract_text()

    # Temukan posisi kalimat "Dengan ini menyatakan sesungguhnya bahwa:"
    match_start = re.search(r"Dengan ini menyatakan sesungguhnya bahwa:", text)

    if match_start:
        # Ambil teks setelah kalimat tersebut
        text_after = text[match_start.end():]

        # Ekstrak Nama setelah "Nama :"
        name_pattern = r"Nama\s*[:\-]?\s*(.*)"  # Ambil seluruh teks setelah "Nama :"
        name_match = re.search(name_pattern, text_after)

        if name_match:
            name = name_match.group(1).strip()
            return name
    return None

# Fungsi untuk mendapatkan nama-nama dari PDF di folder
def get_names_from_folder(folder_path):
    names = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            name = extract_name(pdf_path)
            if name:
                names.append(name)
    return names

# Fungsi untuk mencari nama yang ada di folder B tapi tidak ada di folder A
def find_missing_names(folder_a_path, folder_b_path):
    # Ambil daftar nama dari kedua folder
    names_a = get_names_from_folder(folder_a_path)
    names_b = get_names_from_folder(folder_b_path)

    # Cari nama yang ada di folder B tapi tidak ada di folder A
    missing_names = [name for name in names_b if name not in names_a]

    return missing_names, names_a

# Fungsi untuk mendeteksi nama duplikat di folder A
def find_duplicates_in_folder(names):
    # Hitung kemunculan nama di folder A
    name_counts = Counter(names)
    # Ambil nama yang muncul lebih dari sekali
    duplicates = [name for name, count in name_counts.items() if count > 1]
    return duplicates

# Fungsi untuk menampilkan hasil pencarian
def display_results():
    # Mencari nama yang hilang dari folder A dan mendapatkan daftar nama di folder A
    missing_names, names_a = find_missing_names(folder_a_path, folder_b_path)

    # Menampilkan hasil nama yang hilang
    if missing_names:
        print("\nNama yang ada di Folder B tapi tidak ada di Folder A:")
        for name in missing_names:
            print(name)
    else:
        print("Tidak ada nama yang hilang di Folder A.")
    
    # Mencari dan menampilkan nama duplikat di folder A
    duplicates = find_duplicates_in_folder(names_a)
    if duplicates:
        print("\nNama duplikat di Folder A:")
        for name in duplicates:
            print(name)
    else:
        print("\nTidak ada nama duplikat di Folder A.")

# Menjalankan fungsi pencarian dan deteksi duplikat
display_results()
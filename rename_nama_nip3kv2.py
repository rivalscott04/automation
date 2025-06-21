import os
import re
from PyPDF2 import PdfReader
import shutil

# --- Konfigurasi ---
source_folder_path = r"C:\Users\Rival\Documents\SPMT P3K T1 2025"
renamed_folder_path =  r"C:\Users\Rival\Documents\SPMT P3K T1 2025\renamed"

# --- Fungsi Inti ---

def extract_nip_and_name(pdf_path):
    """Mengekstrak NIP dan Nama dari file PDF dengan pola yang fleksibel."""
    try:
        reader = PdfReader(pdf_path)
        page = reader.pages[0]
        text = page.extract_text() or ""

        anchor_pattern = r"Dengan ini menyatakan (?:dengan )?(?:sesungguhnya,?)?\s*bahwa\s*:"
        match_start = re.search(anchor_pattern, text, re.IGNORECASE | re.DOTALL)

        if match_start:
            text_after = text[match_start.end():]
            pattern = r"Nama\s*[:\-]?\s*(.*?)\s*(?:\d\.\s*)?(?:NIP|Nomor\s*Induk\s*PPPK)\s*[:\-]?\s*([\d\s]{18,})"
            match = re.search(pattern, text_after, re.IGNORECASE | re.DOTALL)

            if match:
                name = match.group(1).strip()
                nip = re.sub(r'\s', '', match.group(2))
                if name.startswith(':'): name = name[1:].strip()
                if nip.isdigit() and len(nip) >= 18:
                    return nip, name
    except Exception as e:
        print(f"\n[ERROR] Gagal membaca file {os.path.basename(pdf_path)}: {e}")
    return None, None

def sanitize_filename(filename):
    """Membersihkan nama file dari karakter ilegal."""
    return re.sub(r'[\\/*?:"<>|\n]', '', filename)

# --- Alur Utama Script ---

# 1. TAHAP PENGECEKAN DUPLIKAT
print("--- TAHAP 1: Memindai Duplikasi Nama File ---")
potential_names = {}
pdf_files = [f for f in os.listdir(source_folder_path) if f.endswith(".pdf")]
total_files = len(pdf_files)

if total_files == 0:
    print("Tidak ada file PDF yang ditemukan di folder sumber.")
    exit()

for i, filename in enumerate(pdf_files, 1):
    # \r (carriage return) membuat baris ditimpa, menciptakan efek loading
    print(f"Memindai file {i}/{total_files}: {filename.ljust(50)}", end='\r')
    pdf_path = os.path.join(source_folder_path, filename)
    nip, name = extract_nip_and_name(pdf_path)
    if nip and name:
        new_name = f"SPMT_{nip}_{name}.pdf"
        sanitized_name = sanitize_filename(new_name)
        # Kumpulkan semua file asli yang akan menghasilkan nama baru yang sama
        if sanitized_name not in potential_names:
            potential_names[sanitized_name] = []
        potential_names[sanitized_name].append(filename)

# Beri baris baru setelah progress selesai
print("\n\nPemindaian selesai.")

# Cari nama file yang akan duplikat
duplicates = {name: files for name, files in potential_names.items() if len(files) > 1}

# 2. TAHAP INTERAKTIF JIKA ADA DUPLIKAT
user_choice = '2' # Default: rename dengan nomor urut
if duplicates:
    print("\n[PERINGATAN] Ditemukan potensi nama file yang sama:")
    for name, files in duplicates.items():
        print(f"- Nama Baru: {name}")
        print(f"  Berasal dari: {', '.join(files)}")
    
    print(f"\nTotal ditemukan: {len(duplicates)} nama file yang akan duplikat.")
    print("\nBagaimana Anda ingin melanjutkan?")
    print("  1. Lanjutkan (Hanya salin file pertama, lewati duplikat)")
    print("  2. Lanjutkan (Tambahkan nomor urut seperti (1), (2) pada file duplikat)")
    print("  3. Batalkan Proses")
    
    while True:
        choice = input("Masukkan pilihan Anda (1/2/3): ").strip()
        if choice in ['1', '2', '3']:
            user_choice = choice
            break
        else:
            print("Pilihan tidak valid, silakan coba lagi.")
else:
    print("Tidak ada potensi duplikasi nama file. Proses akan dilanjutkan secara otomatis.")

if user_choice == '3':
    print("Proses dibatalkan oleh pengguna.")
    exit()

# 3. TAHAP EKSEKUSI
print("\n--- TAHAP 2: Memproses dan Merename File ---")
if not os.path.exists(renamed_folder_path):
    os.makedirs(renamed_folder_path)

processed_names = {} # Digunakan untuk melacak nama yang sudah diproses

for i, filename in enumerate(pdf_files, 1):
    pdf_path = os.path.join(source_folder_path, filename)
    nip, name = extract_nip_and_name(pdf_path)

    if nip and name:
        new_name = f"SPMT_{nip}_{name}.pdf"
        sanitized_name = sanitize_filename(new_name)
        final_path = ""

        if user_choice == '1':
            if sanitized_name in processed_names:
                print(f"[SKIP] Melewatkan duplikat untuk: {filename}")
                continue # Lewati file ini
            processed_names[sanitized_name] = True
            final_path = os.path.join(renamed_folder_path, sanitized_name)
        
        elif user_choice == '2':
            base, ext = os.path.splitext(sanitized_name)
            counter = processed_names.get(base, 0) + 1
            processed_names[base] = counter
            
            if counter == 1:
                final_name = f"{base}{ext}"
            else:
                final_name = f"{base}({counter-1}){ext}"
                # Jika nama dengan (1) sudah ada, cari nama unik berikutnya
                temp_path = os.path.join(renamed_folder_path, final_name)
                if os.path.exists(temp_path):
                    temp_counter = counter - 1
                    while os.path.exists(os.path.join(renamed_folder_path, f"{base}({temp_counter}){ext}")):
                        temp_counter += 1
                    final_name = f"{base}({temp_counter}){ext}"
            
            final_path = os.path.join(renamed_folder_path, final_name)

        if final_path:
            shutil.copy(pdf_path, final_path)
            print(f"[OK] {filename} -> {os.path.basename(final_path)}")
    else:
        print(f"[GAGAL] Tidak bisa mengekstrak data dari: {filename}")

print("\nProses selesai.")
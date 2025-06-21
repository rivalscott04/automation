import os
import re
import shutil
import threading
import customtkinter as ctk
from tkinter import filedialog
from PyPDF2 import PdfReader

# --- Konfigurasi Awal & Tampilan ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF Renamer Pro")
        self.geometry("800x650")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(3, weight=1)
        
        # --- Variabel Internal ---
        self.source_path = ctk.StringVar()
        self.dest_path = ctk.StringVar()
        self.is_running = False
        self.user_choice = "2" # Default choice for duplicates

        # --- Membuat Widget Tampilan ---

        # 1. Folder Sumber
        self.source_label = ctk.CTkLabel(self, text="Folder Sumber:")
        self.source_label.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")
        
        self.source_entry = ctk.CTkEntry(self, textvariable=self.source_path, width=400)
        self.source_entry.grid(row=0, column=1, padx=0, pady=(20, 5), sticky="ew")
        
        self.source_button = ctk.CTkButton(self, text="Browse", command=self.select_source_folder)
        self.source_button.grid(row=0, column=2, padx=20, pady=(20, 5))

        # 2. Folder Tujuan
        self.dest_label = ctk.CTkLabel(self, text="Folder Tujuan:")
        self.dest_label.grid(row=1, column=0, padx=20, pady=5, sticky="w")

        self.dest_entry = ctk.CTkEntry(self, textvariable=self.dest_path, width=400)
        self.dest_entry.grid(row=1, column=1, padx=0, pady=5, sticky="ew")

        self.dest_button = ctk.CTkButton(self, text="Browse", command=self.select_dest_folder)
        self.dest_button.grid(row=1, column=2, padx=20, pady=5)

        # 3. Tombol Start
        self.start_button = ctk.CTkButton(self, text="Mulai Proses", command=self.start_processing_thread, height=40)
        self.start_button.grid(row=2, column=0, columnspan=3, padx=20, pady=20, sticky="ew")

        # 4. Log Proses
        self.log_textbox = ctk.CTkTextbox(self, state="disabled", wrap="word")
        self.log_textbox.grid(row=3, column=0, columnspan=3, padx=20, pady=(0, 5), sticky="nsew")
        
        # 5. Progress Bar
        self.progressbar = ctk.CTkProgressBar(self)
        self.progressbar.set(0)
        self.progressbar.grid(row=4, column=0, columnspan=3, padx=20, pady=(0, 20), sticky="ew")
        
    # --- Fungsi untuk Interaksi GUI ---

    def log(self, message):
        """Menambahkan pesan ke log textbox dengan aman dari thread manapun."""
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end")

    def select_source_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.source_path.set(path)
            if not self.dest_path.get():
                self.dest_path.set(os.path.join(path, "renamed"))

    def select_dest_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.dest_path.set(path)
            
    def start_processing_thread(self):
        """Memulai proses di thread terpisah agar GUI tidak freeze."""
        if self.is_running:
            return
        
        if not self.source_path.get() or not self.dest_path.get():
            self.log("[ERROR] Harap pilih Folder Sumber dan Tujuan terlebih dahulu.")
            return

        self.is_running = True
        self.start_button.configure(state="disabled", text="Sedang Memproses...")
        self.progressbar.set(0)
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        
        thread = threading.Thread(target=self.run_process)
        thread.start()

    # --- Fungsi Dialog & Jendela Baru ---

    def show_duplicate_list_window(self, duplicates):
        """[BARU] Menampilkan jendela baru berisi daftar file duplikat."""
        list_win = ctk.CTkToplevel(self)
        list_win.title("Daftar Potensi Duplikat")
        list_win.geometry("750x500")
        list_win.transient(self)
        list_win.grab_set()

        textbox = ctk.CTkTextbox(list_win, wrap="word", font=("Consolas", 12))
        textbox.pack(padx=20, pady=20, fill="both", expand=True)

        header = "Berikut adalah daftar file asli yang akan menghasilkan nama baru yang sama:\n\n"
        textbox.insert("end", header)
        
        for new_name, original_files in duplicates.items():
            entry = f"Nama Baru: {new_name}\n"
            entry += f"  - Berasal dari: {', '.join(original_files)}\n\n"
            textbox.insert("end", entry)
            
        textbox.configure(state="disabled")

    def ask_duplicate_choice_dialog(self, duplicates):
        """Menampilkan dialog untuk menanyakan pilihan user."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Duplikasi Ditemukan")
        dialog.geometry("650x150")
        dialog.transient(self)
        dialog.grab_set()
        
        message = f"Ditemukan {len(duplicates)} potensi nama file yang sama.\nBagaimana Anda ingin melanjutkan?"
        label = ctk.CTkLabel(dialog, text=message, wraplength=450)
        label.pack(padx=20, pady=20)
        
        def set_choice(choice):
            self.user_choice = choice
            dialog.destroy()

        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=10)

        # [BARU] Tombol untuk melihat daftar detail
        view_button = ctk.CTkButton(button_frame, text="Lihat Daftar Duplikat", command=lambda: self.show_duplicate_list_window(duplicates), fg_color="#555555", hover_color="#444444")
        view_button.pack(side="left", padx=10)

        choice1_button = ctk.CTkButton(button_frame, text="Lewati Duplikat", command=lambda: set_choice("1"))
        choice1_button.pack(side="left", padx=10)

        choice2_button = ctk.CTkButton(button_frame, text="Beri Nomor Urut", command=lambda: set_choice("2"))
        choice2_button.pack(side="left", padx=10)

        choice3_button = ctk.CTkButton(button_frame, text="Batalkan", command=lambda: set_choice("3"), fg_color="#E53935", hover_color="#C62828")
        choice3_button.pack(side="left", padx=10)
        
        self.wait_window(dialog)

    # --- Fungsi Logika Inti (dijalankan di thread terpisah) ---
    
    def extract_nip_and_name(self, pdf_path):
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
            self.log(f"[ERROR] Gagal membaca file {os.path.basename(pdf_path)}: {e}")
        return None, None

    def sanitize_filename(self, filename):
        return re.sub(r'[\\/*?:"<>|\n]', '', filename)

    def run_process(self):
        """Fungsi utama yang menjalankan seluruh logika renaming."""
        try:
            source_folder = self.source_path.get()
            dest_folder = self.dest_path.get()

            # TAHAP 1: PEMINDAIAN DUPLIKAT
            self.log("--- TAHAP 1: Memindai Duplikasi Nama File ---")
            potential_names = {}
            pdf_files = [f for f in os.listdir(source_folder) if f.endswith(".pdf")]
            total_files = len(pdf_files)
            if total_files == 0:
                self.log("Tidak ada file PDF yang ditemukan di folder sumber.")
                return

            for i, filename in enumerate(pdf_files, 1):
                self.log(f"Memindai file {i}/{total_files}: {filename}")
                self.progressbar.set(i / total_files)
                pdf_path = os.path.join(source_folder, filename)
                nip, name = self.extract_nip_and_name(pdf_path)
                if nip and name:
                    new_name = f"SPMT_{nip}_{name}.pdf"
                    sanitized_name = self.sanitize_filename(new_name)
                    if sanitized_name not in potential_names:
                        potential_names[sanitized_name] = []
                    potential_names[sanitized_name].append(filename)

            self.log("\nPemindaian selesai.")
            duplicates = {name: files for name, files in potential_names.items() if len(files) > 1}

            # TAHAP 2: INTERAKTIF JIKA ADA DUPLIKAT
            if duplicates:
                self.log("\n[PERINGATAN] Ditemukan potensi nama file yang sama. Menunggu pilihan pengguna...")
                self.ask_duplicate_choice_dialog(duplicates)
            
            if self.user_choice == '3':
                self.log("Proses dibatalkan oleh pengguna.")
                return

            # TAHAP 3: EKSEKUSI
            self.log("\n--- TAHAP 2: Memproses dan Merename File ---")
            if not os.path.exists(dest_folder):
                os.makedirs(dest_folder)

            processed_names = {}
            for i, filename in enumerate(pdf_files, 1):
                self.progressbar.set(i / total_files)
                pdf_path = os.path.join(source_folder, filename)
                nip, name = self.extract_nip_and_name(pdf_path)
                if nip and name:
                    new_name = f"SPMT_{nip}_{name}.pdf"
                    sanitized_name = self.sanitize_filename(new_name)
                    final_path = ""
                    if self.user_choice == '1':
                        if sanitized_name in processed_names:
                            self.log(f"[SKIP] Melewatkan duplikat untuk: {filename}")
                            continue
                        processed_names[sanitized_name] = True
                        final_path = os.path.join(dest_folder, sanitized_name)
                    elif self.user_choice == '2':
                        base, ext = os.path.splitext(sanitized_name)
                        counter = processed_names.get(base, 0) + 1
                        processed_names[base] = counter
                        final_name = f"{base}{ext}" if counter == 1 else f"{base}({counter-1}){ext}"
                        final_path = os.path.join(dest_folder, final_name)
                        while os.path.exists(final_path):
                            final_name = f"{base}({counter}){ext}"
                            final_path = os.path.join(dest_folder, final_name)
                            counter += 1
                    if final_path:
                        shutil.copy(pdf_path, final_path)
                        self.log(f"[OK] {filename} -> {os.path.basename(final_path)}")
                else:
                    self.log(f"[GAGAL] Tidak bisa mengekstrak data dari: {filename}")
            self.log("\nProses selesai.")
        finally:
            self.is_running = False
            self.start_button.configure(state="normal", text="Mulai Proses")

if __name__ == "__main__":
    app = App()
    app.mainloop()
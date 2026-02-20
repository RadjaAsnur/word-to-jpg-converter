import customtkinter as ctk
from tkinter import filedialog, messagebox
from docx2pdf import convert
from pdf2image import convert_from_path
import os
import threading
import sys

# Kode untuk mencegah error 'NoneType' pada aplikasi tanpa konsol
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

# Konfigurasi Tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class WordToJpgApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Word to JPG Converter - KingAsnur")
        self.geometry("500x450")

# --- MENAMBAHKAN LOGO JENDELA ---
        try:
            # Gunakan file .ico yang sudah kamu siapkan
            self.iconbitmap("logo.ico") 
        except:
            pass # Agar aplikasi tidak error jika file logo tidak ditemukan

        # --- UI ELEMENTS ---
        self.label = ctk.CTkLabel(self, text="Word to JPG Converter", font=("Roboto", 24, "bold"))
        self.label.pack(pady=(30, 10))

        # Pilih File
        self.btn_file = ctk.CTkButton(self, text="1. Pilih File Word (.docx)", command=self.select_file)
        self.btn_file.pack(pady=10)
        self.file_label = ctk.CTkLabel(self, text="File belum dipilih", font=("Roboto", 11), text_color="gray")
        self.file_label.pack()

        # Pilih Folder
        self.btn_folder = ctk.CTkButton(self, text="2. Pilih Folder Tujuan", command=self.select_folder, fg_color="#444444")
        self.btn_folder.pack(pady=10)
        self.folder_label = ctk.CTkLabel(self, text="Folder belum dipilih", font=("Roboto", 11), text_color="gray")
        self.folder_label.pack()

        # Tombol Konversi
        self.btn_convert = ctk.CTkButton(
            self, text="Konversi Sekarang", 
            command=self.start_conversion_thread, 
            state="disabled", 
            fg_color="#2ecc71", 
            hover_color="#27ae60"
        )
        self.btn_convert.pack(pady=30)

        self.status_label = ctk.CTkLabel(self, text="", text_color="yellow")
        self.status_label.pack()

        # --- WATERMARK ---
        self.watermark = ctk.CTkLabel(
            self, 
            text="Created by KingAsnur", 
            font=("Roboto", 10, "italic"), 
            text_color="gray50"
        )
        self.watermark.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-10)

        # Variables
        self.file_path = ""
        self.output_path = ""

    # --- LOGIKA TOMBOL ---
    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if self.file_path:
            self.file_label.configure(text=f"File: {os.path.basename(self.file_path)}")
            self.check_ready()

    def select_folder(self):
        self.output_path = filedialog.askdirectory()
        if self.output_path:
            self.folder_label.configure(text=f"Tujuan: {self.output_path}")
            self.check_ready()

    def check_ready(self):
        if self.file_path and self.output_path:
            self.btn_convert.configure(state="normal")

    def start_conversion_thread(self):
        # Daemon=True agar thread berhenti otomatis jika aplikasi ditutup
        threading.Thread(target=self.convert_process, daemon=True).start()

    # --- FUNGSI UTAMA (OTOMATIS) ---
    def convert_process(self):
        try:
            self.btn_convert.configure(state="disabled")
            self.status_label.configure(text="Sedang memproses... Harap tunggu.")

            # 1. Tentukan Base Path (Penting untuk EXE & Portabilitas)
            if getattr(sys, 'frozen', False):
                # Jika berjalan sebagai .exe
                base_path = sys._MEIPASS
            else:
                # Jika berjalan sebagai skrip .py
                base_path = os.path.dirname(os.path.abspath(__file__))

            # 2. Path Otomatis ke Poppler Bin
            poppler_bin = os.path.join(base_path, 'poppler', 'Library', 'bin')
            
            # File temporary PDF
            temp_pdf = os.path.join(self.output_path, "temp_result_hidden.pdf")

            # 3. Konversi Word -> PDF
            # (Memerlukan MS Word terinstal di Windows)
            convert(self.file_path, temp_pdf)

            # 4. Konversi PDF -> JPG menggunakan Poppler Otomatis
            images = convert_from_path(temp_pdf, poppler_path=poppler_bin)

            # Simpan setiap halaman
            for i, image in enumerate(images):
                save_name = f"{os.path.splitext(os.path.basename(self.file_path))[0]}_hal_{i+1}.jpg"
                image.save(os.path.join(self.output_path, save_name), "JPEG")

            # 5. Bersihkan Temp PDF
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)

            self.status_label.configure(text="Konversi Berhasil!")
            messagebox.showinfo("Sukses", f"Halo KingAsnur, file JPG kamu sudah siap di folder tujuan!")

        except Exception as e:
            self.status_label.configure(text="Gagal.")
            messagebox.showerror("Error", f"Terjadi kendala: {str(e)}")
        finally:
            self.btn_convert.configure(state="normal")

if __name__ == "__main__":
    app = WordToJpgApp()
    app.mainloop()
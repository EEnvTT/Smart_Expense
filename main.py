import tkinter as tk
import pandas as pd
import os
from fpdf import FPDF
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
import csv

# --- 1. Inisialisasi Database ---
def init_db():
    conn = sqlite3.connect('smart_expense_v2.db')
    c = conn.cursor()
    # Menambahkan kolom tanggal dan kategori
    c.execute('''CREATE TABLE IF NOT EXISTS transaksi
                 (id INTEGER PRIMARY KEY, tanggal TEXT, tipe TEXT, kategori TEXT, jumlah INTEGER, keterangan TEXT)''')
    conn.commit()
    conn.close()

# --- 2. Fungsi Logika Aplikasi ---
def update_kategori(*args):
    # Mengubah pilihan kategori berdasarkan tipe (Pemasukan/Pengeluaran)
    if var_tipe.get() == "Pengeluaran":
        combo_kategori['values'] = ("Makan & Minum", "Bayar Kost/Tagihan", "Transportasi", "Kebutuhan Kuliah", "Hiburan", "Lainnya")
    else:
        combo_kategori['values'] = ("Kiriman Orang Tua", "Gaji/Freelance", "Beasiswa", "Lainnya")
    combo_kategori.current(0)

def tambah_data():
    tipe = var_tipe.get()
    kategori = var_kategori.get()
    jumlah = entry_jumlah.get()
    ket = entry_ket.get()
    tanggal = datetime.now().strftime("%Y-%m-%d %H:%M") # Format waktu saat ini

    if not jumlah.isdigit():
        messagebox.showerror("Error", "Masukkan jumlah dalam bentuk angka (tanpa titik/koma)!")
        return
    if not ket:
        messagebox.showerror("Error", "Keterangan tidak boleh kosong!")
        return

    conn = sqlite3.connect('smart_expense_v2.db')
    c = conn.cursor()
    c.execute("INSERT INTO transaksi (tanggal, tipe, kategori, jumlah, keterangan) VALUES (?, ?, ?, ?, ?)", 
              (tanggal, tipe, kategori, int(jumlah), ket))
    conn.commit()
    conn.close()
    
    entry_jumlah.delete(0, tk.END)
    entry_ket.delete(0, tk.END)
    muat_data()
    messagebox.showinfo("Sukses", "Data berhasil ditambahkan!")

def hapus_data():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Peringatan", "Pilih data di tabel terlebih dahulu untuk dihapus!")
        return
    
    konfirmasi = messagebox.askyesno("Konfirmasi Hapus", "Apakah Anda yakin ingin menghapus data ini?")
    if konfirmasi:
        # Mengambil ID dari baris yang dipilih (ID ada di kolom index 0 yang disembunyikan)
        item_id = tree.item(selected_item)['values'][0]
        
        conn = sqlite3.connect('smart_expense_v2.db')
        c = conn.cursor()
        c.execute("DELETE FROM transaksi WHERE id=?", (item_id,))
        conn.commit()
        conn.close()
        
        muat_data()

def export_csv():
    conn = sqlite3.connect('smart_expense_v2.db')
    c = conn.cursor()
    c.execute("SELECT tanggal, tipe, kategori, jumlah, keterangan FROM transaksi")
    rows = c.fetchall()
    conn.close()

    if not rows:
        messagebox.showwarning("Peringatan", "Tidak ada data untuk diexport!")
        return

    # Membuka dialog untuk menyimpan file
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", 
                                             filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                                             title="Simpan Laporan Keuangan")
    if file_path:
        try:
            with open(file_path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["Tanggal", "Tipe", "Kategori", "Jumlah", "Keterangan"]) # Header
                writer.writerows(rows)
            messagebox.showinfo("Sukses", f"Data berhasil diexport ke:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengeksport data:\n{str(e)}")

def import_data():
    # Buka dialog untuk memilih file CSV atau Excel
    file_path = filedialog.askopenfilename(
        title="Pilih File untuk Diimport",
        filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
    )
    
    if not file_path:
        return # Batal memilih file
        
    try:
        # Deteksi ekstensi file untuk menentukan cara membacanya
        _, ext = os.path.splitext(file_path)
        
        if ext.lower() == '.csv':
            df = pd.read_csv(file_path)
        elif ext.lower() in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        else:
            messagebox.showerror("Error", "Format file tidak didukung! Gunakan CSV atau Excel.")
            return
            
        # Validasi kolom (Pastikan header file sesuai standar kita)
        expected_columns = ["Tanggal", "Tipe", "Kategori", "Jumlah", "Keterangan"]
        if not all(col in df.columns for col in expected_columns):
            messagebox.showerror("Error", f"Format kolom tidak sesuai!\nPastikan baris pertama file memiliki kolom:\n{expected_columns}")
            return
            
        conn = sqlite3.connect('smart_expense_v2.db')
        c = conn.cursor()
        
        # Loop data dari file dan masukkan ke database
        berhasil = 0
        for index, row in df.iterrows():
            try:
                # Memastikan jumlah berupa angka sebelum dimasukkan
                jumlah = int(row["Jumlah"])
                
                # Menangani nilai kosong (NaN) agar menjadi string kosong
                keterangan = str(row["Keterangan"]) if pd.notna(row["Keterangan"]) else ""
                
                c.execute("INSERT INTO transaksi (tanggal, tipe, kategori, jumlah, keterangan) VALUES (?, ?, ?, ?, ?)", 
                          (str(row["Tanggal"]), str(row["Tipe"]), str(row["Kategori"]), jumlah, keterangan))
                berhasil += 1
            except ValueError:
                # Lewati baris jika data 'Jumlah' bukan angka (mencegah aplikasi error)
                continue 
                
        conn.commit()
        conn.close()
        
        muat_data() # Refresh tabel UI
        messagebox.showinfo("Sukses", f"Berhasil mengimport {berhasil} baris data!")
        
    except Exception as e:
        messagebox.showerror("Error", f"Gagal mengimport data:\n{str(e)}")

def export_pdf():
    # Ambil data dari database
    conn = sqlite3.connect('smart_expense_v2.db')
    c = conn.cursor()
    c.execute("SELECT tanggal, tipe, kategori, jumlah, keterangan FROM transaksi ORDER BY tanggal ASC")
    rows = c.fetchall()
    conn.close()

    if not rows:
        messagebox.showwarning("Peringatan", "Tidak ada data untuk diexport!")
        return

    # Buka dialog penyimpanan file
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                             filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                                             title="Simpan Laporan PDF")
    if not file_path:
        return # Batal simpan

    try:
        # Inisialisasi PDF
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.add_page()
        
        # Judul Laporan
        pdf.set_font("Arial", 'B', 16)
        pdf.set_text_color(0, 0, 128) # Navy blue
        pdf.cell(190, 10, "Laporan Keuangan Smart Expense", ln=True, align='C')
        pdf.ln(5)

        # Warna Header Tabel (Gradasi/Campuran Biru dan Ungu - SlateBlue)
        pdf.set_fill_color(106, 90, 205) 
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", 'B', 10)

        # Pengaturan Lebar Kolom
        col_widths = [35, 25, 35, 30, 65]
        headers = ["Tanggal", "Tipe", "Kategori", "Jumlah (Rp)", "Keterangan"]

        # Cetak Header Tabel
        for i in range(len(headers)):
            pdf.cell(col_widths[i], 10, headers[i], border=1, align='C', fill=True)
        pdf.ln()

        # Isi Tabel
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", '', 9)
        
        fill = False
        total_saldo = 0
        
        for row in rows:
            tanggal, tipe, kategori, jumlah, keterangan = row
            
            # Kalkulasi Saldo Akhir
            if tipe == "Pemasukan":
                total_saldo += jumlah
            else:
                total_saldo -= jumlah

            # Mewarnai baris secara selang-seling (Zebra striping) dengan warna biru keputihan
            if fill:
                pdf.set_fill_color(240, 248, 255) # Alice blue
            else:
                pdf.set_fill_color(255, 255, 255) # Putih

            pdf.cell(col_widths[0], 8, str(tanggal)[:10], border=1, align='C', fill=fill)
            pdf.cell(col_widths[1], 8, str(tipe), border=1, align='C', fill=fill)
            pdf.cell(col_widths[2], 8, str(kategori), border=1, align='L', fill=fill)
            pdf.cell(col_widths[3], 8, f"{jumlah:,}", border=1, align='R', fill=fill)
            pdf.cell(col_widths[4], 8, str(keterangan), border=1, align='L', fill=fill)
            pdf.ln()
            
            fill = not fill # Ubah status fill untuk baris berikutnya

        pdf.ln(5)
        
        # Cetak Total Saldo Akhir
        pdf.set_font("Arial", 'B', 12)
        if total_saldo < 0:
            pdf.set_text_color(255, 0, 0) # Merah jika minus
        else:
            pdf.set_text_color(0, 128, 0) # Hijau jika plus
            
        pdf.cell(190, 10, f"Total Saldo Akhir: Rp {total_saldo:,}", ln=True, align='R')

        # Simpan File
        pdf.output(file_path)
        messagebox.showinfo("Sukses", f"Data berhasil diexport ke PDF:\n{file_path}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Gagal mengeksport data PDF:\n{str(e)}")

def muat_data():
    for row in tree.get_children():
        tree.delete(row)
        
    conn = sqlite3.connect('smart_expense_v2.db')
    c = conn.cursor()
    c.execute("SELECT * FROM transaksi ORDER BY tanggal DESC") # Tampilkan yang terbaru di atas
    rows = c.fetchall()
    
    total_saldo = 0
    for row in rows:
        # Insert ke treeview (row[0] adalah ID yang akan disembunyikan)
        tree.insert("", tk.END, values=(row[0], row[1], row[2], row[3], f"Rp {row[4]:,}", row[5]))
        if row[2] == "Pemasukan":
            total_saldo += row[4]
        else:
            total_saldo -= row[4]
            
    label_saldo.config(text=f"Total Saldo: Rp {total_saldo:,}")
    
    # Merubah warna teks saldo berdasarkan kondisi keuangan
    if total_saldo < 0:
        label_saldo.config(fg="red")
    else:
        label_saldo.config(fg="#008000") # Hijau
        
    conn.close()

# --- 3. Setup Antarmuka (GUI) ---
root = tk.Tk()
root.title("Smart Expense V2")
root.geometry("750x650") # Diperlebar agar tabel muat
root.configure(bg="#f0f8ff")

label_judul = tk.Label(root, text="Smart Expense - Manajemen Keuangan Kost", font=("Helvetica", 16, "bold"), bg="#f0f8ff", fg="#000080")
label_judul.pack(pady=15)

# --- Frame Input ---
frame_input = tk.LabelFrame(root, text="Input Transaksi Baru", bg="#ffffff", padx=10, pady=10, font=("Helvetica", 10, "bold"))
frame_input.pack(pady=5, padx=20, fill="x")

# Tipe
tk.Label(frame_input, text="Tipe:", bg="#ffffff").grid(row=0, column=0, sticky="w", pady=5)
var_tipe = tk.StringVar(value="Pengeluaran")
var_tipe.trace("w", update_kategori) # Trigger update_kategori jika tipe berubah
tk.Radiobutton(frame_input, text="Pengeluaran", variable=var_tipe, value="Pengeluaran", bg="#ffffff").grid(row=0, column=1, sticky="w")
tk.Radiobutton(frame_input, text="Pemasukan", variable=var_tipe, value="Pemasukan", bg="#ffffff").grid(row=0, column=2, sticky="w")

# Kategori
tk.Label(frame_input, text="Kategori:", bg="#ffffff").grid(row=1, column=0, sticky="w", pady=5)
var_kategori = tk.StringVar()
combo_kategori = ttk.Combobox(frame_input, textvariable=var_kategori, state="readonly", width=27)
combo_kategori.grid(row=1, column=1, columnspan=2, sticky="w", pady=5)
update_kategori() # Set nilai awal combobox

# Jumlah
tk.Label(frame_input, text="Jumlah (Rp):", bg="#ffffff").grid(row=2, column=0, sticky="w", pady=5)
entry_jumlah = tk.Entry(frame_input, width=30)
entry_jumlah.grid(row=2, column=1, columnspan=2, sticky="w", pady=5)

# Keterangan
tk.Label(frame_input, text="Keterangan:", bg="#ffffff").grid(row=3, column=0, sticky="w", pady=5)
entry_ket = tk.Entry(frame_input, width=45)
entry_ket.grid(row=3, column=1, columnspan=3, sticky="w", pady=5)

# Tombol Tambah
btn_tambah = tk.Button(frame_input, text="➕ Tambah Transaksi", bg="#4169e1", fg="white", font=("Helvetica", 10, "bold"), command=tambah_data)
btn_tambah.grid(row=4, column=0, columnspan=4, pady=10, sticky="ew")

# --- Frame Saldo & Aksi ---
frame_aksi = tk.Frame(root, bg="#f0f8ff")
frame_aksi.pack(fill="x", padx=20, pady=5)

label_saldo = tk.Label(frame_aksi, text="Total Saldo: Rp 0", font=("Helvetica", 14, "bold"), bg="#f0f8ff", fg="#008000")
label_saldo.pack(side="left")

btn_import = tk.Button(frame_aksi, text="📥 Import Excel/CSV", bg="#ff8c00", fg="white", command=import_data)
btn_import.pack(side="left", padx=15)

btn_pdf = tk.Button(frame_aksi, text="📕 Export ke PDF", bg="#b22222", fg="white", command=export_pdf)
btn_pdf.pack(side="right", padx=5)

btn_export = tk.Button(frame_aksi, text="📄 Export ke CSV", bg="#20b2aa", fg="white", command=export_csv)
btn_export.pack(side="right", padx=5)

btn_hapus = tk.Button(frame_aksi, text="🗑️ Hapus Data", bg="#dc143c", fg="white", command=hapus_data)
btn_hapus.pack(side="right", padx=5)

# --- Tabel Riwayat (Treeview) ---
kolom = ("ID", "Tanggal", "Tipe", "Kategori", "Jumlah", "Keterangan")
tree = ttk.Treeview(root, columns=kolom, show="headings", height=12)

# Mengatur heading
tree.heading("ID", text="ID")
tree.heading("Tanggal", text="Tanggal")
tree.heading("Tipe", text="Tipe")
tree.heading("Kategori", text="Kategori")
tree.heading("Jumlah", text="Jumlah")
tree.heading("Keterangan", text="Keterangan")

# Mengatur kolom (Menyembunyikan kolom ID)
tree.column("ID", width=0, stretch=tk.NO) 
tree.column("Tanggal", width=120, anchor="center")
tree.column("Tipe", width=100, anchor="center")
tree.column("Kategori", width=120, anchor="center")
tree.column("Jumlah", width=120, anchor="e")
tree.column("Keterangan", width=200, anchor="w")

tree.pack(pady=10, padx=20, fill="both", expand=True)

# Jalankan database & GUI
init_db()
muat_data()
root.mainloop()
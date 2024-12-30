import pandas as pd
from tkinter import filedialog
import tkinter as tk

# Function to read Excel file with multiple sheets
def baca_file_excel(nama_file):
    try:
        sheets = pd.read_excel(nama_file, sheet_name=None)  # Read all sheets
        print("\nData berhasil dimuat dari file Excel!")
        return sheets
    except Exception as e:
        print(f"Terjadi kesalahan saat membaca file: {e}")
        return None

# Function to search data based on criteria
def cari_data(data, kriteria):
    print("\nKolom yang tersedia dalam sheet:")
    print(data.columns.tolist())  # Display available columns

    # Remove empty or irrelevant columns
    data = data.dropna(how='all', axis=1)
    data.columns = data.columns.str.strip()  # Normalize column names

    hasil = data.copy()
    for key, value in kriteria.items():
        if value:
            if key in hasil.columns:  # Check if column exists
                hasil = hasil[hasil[key].astype(str).str.contains(value, case=False, na=False)]
            else:
                print(f"Kolom '{key}' tidak ditemukan dalam data!")
                return pd.DataFrame()  # Return empty DataFrame if column not found
    return hasil

# Function to display data
def tampilkan_data(data):
    if data.empty:
        print("\nTidak ada data yang sesuai!")
    else:
        print(data)

# Function to select a file
def pilih_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Pilih file Excel",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return file_path

# Main program
def main():
    print("=== Sistem Pengukur Nilai Tukar Petani (NTP) ===")
    print("\nSilakan pilih file Excel...")
    nama_file = pilih_file()
    if not nama_file:  # If the user cancels the file selection
        print("Pemilihan file dibatalkan")
        return
        
    sheets = baca_file_excel(nama_file)
    if sheets is None:
        return

    while True:
        print("\n=== Menu Utama ===")
        print("1. Tampilkan Data")
        print("2. Cari Data")
        print("3. Keluar")
        pilihan = input("Pilih menu (1/2/3): ")

        if pilihan == "1":
            print("\nSheet yang Tersedia:")
            for i, sheet in enumerate(sheets.keys(), start=1):
                print(f"{i}. {sheet}")
            pilih_sheet = input("Pilih sheet yang ingin ditampilkan (1/2): ")

            try:
                sheet_name = list(sheets.keys())[int(pilih_sheet) - 1]
                print(f"\nMenampilkan data dari sheet: {sheet_name}")
                tampilkan_data(sheets[sheet_name])
            except (IndexError, ValueError):
                print("Pilihan sheet tidak valid!")

        elif pilihan == "2":
            print("\nSheet yang Tersedia:")
            for i, sheet in enumerate(sheets.keys(), start=1):
                print(f"{i}. {sheet}")
            pilih_sheet = input("Pilih sheet untuk mencari data (1/2): ")

            try:
                sheet_name = list(sheets.keys())[int(pilih_sheet) - 1]
                data = sheets[sheet_name]
            except (IndexError, ValueError):
                print("Pilihan sheet tidak valid!")
                continue

            print("\nMasukkan Kriteria Pencarian (kosongkan jika tidak diperlukan):")
            wilayah = input("Wilayah: ")
            bulan = input("Bulan: ")

            kriteria = {
                "Wilayah": wilayah,
                "Bulan": bulan,
            }

            hasil = cari_data(data, kriteria)
            print(f"\nHasil Pencarian di sheet: {sheet_name}")
            tampilkan_data(hasil)

        elif pilihan == "3":
            print("Keluar dari program. Terima kasih!")
            break

        else:
            print("Pilihan tidak valid, coba lagi!")

if __name__ == "__main__":
    main()

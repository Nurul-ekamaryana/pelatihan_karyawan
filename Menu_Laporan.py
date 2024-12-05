import openpyxl
import os

FILE_NAME = "Laporan_Pelatihan.xlsx"

# Fungsi untuk menyimpan data ke Excel
def simpan_ke_excel(nama, kehadiran, nilai, status_kelulusan):
    # Jika file Excel tidak ada, buat file baru
    if not os.path.exists(FILE_NAME):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Laporan"
        # Header kolom
        sheet.append(['Nama', 'Kehadiran', 'Nilai', 'Status Kelulusan'])
    else:
        workbook = openpyxl.load_workbook(FILE_NAME)
        sheet = workbook['Laporan']

    # Tambahkan data baru
    sheet.append([nama, kehadiran, nilai, status_kelulusan])

    # Simpan perubahan
    workbook.save(FILE_NAME)
    
def laporan():
    print("\nMAU MELAKUKAN APA DIniali???")
    print("1. Tambah Status")
    print("2. Lihat Satus")
    print("5. Kembali")
    pilihan = input("Pilih Menu: ")
    
    if pilihan == "1":
        masukkan_data()
    elif pilihan == "2":
        lihat_laporan()  
    elif pilihan == "5":
        from main import menu_utama
        menu_utama()
    else:
        print("PILIHAN SALAH!!😒😒")

# Fungsi untuk menentukan status kelulusan
def hitung_status_kehadiran_dan_kelulusan(kehadiran, nilai):
    # Tentukan status kelulusan berdasarkan kehadiran dan nilai
    if kehadiran.lower() == "hadir":
        if nilai > 84:
            return "Lulus"
        else:
            return "Tidak Lulus"
    elif kehadiran.lower() == "tidak hadir":
        return "Tidak Lulus"
    else:
        return "Input Kehadiran Tidak Valid"

# Fungsi untuk memasukkan data pelatihan
def masukkan_data():
    print("\n=== Masukkan Data Kehadiran dan Nilai ===")
    nama = input("Nama Peserta: ")
    kehadiran = input("Status Kehadiran (Hadir/Tidak Hadir): ")
    try:
        nilai = float(input("Nilai Pelatihan: "))
    except ValueError:
        print("Nilai harus berupa angka.")
        return

    # Hitung status kelulusan
    status_kelulusan = hitung_status_kehadiran_dan_kelulusan(kehadiran, nilai)

    # Simpan data ke Excel
    try:
        simpan_ke_excel(nama, kehadiran, nilai, status_kelulusan)
        print(f"Data berhasil disimpan. Status kelulusan: {status_kelulusan}")
        laporan()
    except Exception as e:
        print(f"Terjadi kesalahan saat menyimpan data: {e}")

# Fungsi untuk menampilkan laporan pelatihan
def lihat_laporan():
    # lihat status selesai dan belum selesai
    # lama jadwal pelatihan namabah row tanggal
    #relasi!!!!!!!!!!!!!!!!!
    if not os.path.exists(FILE_NAME):
        print("Laporan tidak ditemukan.")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)
    sheet = workbook['Laporan']
    
    if sheet.max_row == 1:
        print("Belum ada data dalam laporan.")
        workbook.close()
        return

    print("\n=== Laporan Kehadiran dan Hasil Pelatihan ===")
    for row in sheet.iter_rows(min_row=2, values_only=True):
        print(f"Nama: {row[0]}, Kehadiran: {row[1]}, Nilai: {row[2]}, Status Kelulusan: {row[3]}")

    workbook.close()
    laporan()
    



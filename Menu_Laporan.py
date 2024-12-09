import openpyxl
import os

FILE_NAME = "Laporan_Pelatihan.xlsx"

def baca_nilai():
    """Membaca data nilai dari file Excel dan mengembalikan daftar nama."""
    if not os.path.exists("Menu_nilai.xlsx"):
        return []  # Jika file tidak ada, kembalikan list kosong

    workbook = openpyxl.load_workbook("Menu_nilai.xlsx")
    sheet = workbook["nilai"]

    nama_karyawan_list = []
    nama_kegiatan_list = []
    nilai_list = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        nama_karyawan_list.append(row[0])
        nama_kegiatan_list.append(row[1])
        nilai_list.append(row[2])
        # Ambil nama karyawan dari kolom kedua

    return nama_karyawan_list,nama_karyawan_list,nilai_list


# Fungsi untuk menyimpan data ke Excel
def simpan_ke_excel(nama_karyawan, Nama_Kegiatan, nilai, kehadiran, status_kelulusan):
    # Jika file Excel tidak ada, buat file baru
    if not os.path.exists(FILE_NAME):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Laporan"
        # Header kolom
        sheet.append(['nama_karyawan', 'Nama_Kegiatan', 'nilai, kehadiran', 'status_kelulusan'])
    else:
        workbook = openpyxl.load_workbook(FILE_NAME)
        sheet = workbook['Laporan']
    # Tambahkan data baru
        sheet.append([nama_karyawan, Nama_Kegiatan, nilai, kehadiran, status_kelulusan])
    # Simpan perubahan
        workbook.save(FILE_NAME)
    
def laporan():
    print("\nMAU MELAKUKAN APA DIniali???")
    print("1. Tambah Data")
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

 # Baca daftar nama karyawan dari file peserta
    nama_karyawan_list = baca_nilai()
    nama_kegiatan_list = baca_nilai()
    nilai_list = baca_nilai()


    # Tampilkan pilihan nama karyawan
    print("Pilih nama karyawan:")
    for i, nama in enumerate(nama_karyawan_list, start=1):
        print(f"{i}. {nama}")
    print("0. Masukkan nama karyawan baru")
    
    pilihan_nama = int(input("Masukkan pilihan: "))
    if pilihan_nama == 0:
        Nama_karyawan = input("Masukkan Nama Karyawan: ")
    else:
        Nama_karyawan = nama_karyawan_list[pilihan_nama - 1]

 # Tampilkan pilihan nama karyawan
    print("Pilih nama kegiatan:")
    for i, nama_kegiatan in enumerate(nama_kegiatan_list, start=1):
        print(f"{i}. {nama_kegiatan}")
    print("0. Masukkan nama kegiatan baru")
    
    pilih_kegiatan = int(input("Masukkan pilihan: "))
    if pilih_kegiatan == 0:
        Nama_kegiatan = input("Masukkan Nama Kegiatan: ")
    else:
        Nama_kegiatan = nama_kegiatan_list[pilih_kegiatan - 1]

    
 # Tampilkan pilihan nama karyawan
    print("Pilih nama nilai:")
    for i, nilai in enumerate(nama_kegiatan_list, start=1):
        print(f"{i}. {nilai}")
    print("0. Masukkan nilai baru")

    pilih_nilai = int(input("Masukkan pilihan: "))
    if pilih_nilai == 0:
        niali = input("Masukkan Nama Kegiatan: ")
    else:
        niali = nilai_list[pilih_nilai - 1]
   

        kehadiran = input("masukan kehadiran :")
        status_kelulusan = hitung_status_kehadiran_dan_kelulusan(kehadiran, nilai)


    try:
        simpan_ke_excel(Nama_karyawan, nama_kegiatan, nilai, kehadiran, status_kelulusan)
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
        print(f"Nama karyawan: {row[0]}, Nama kegiatan: {row[1]}, Nilai: {row[2]}, Kehadiran: {row[3]} Status Kelulusan: {row[4]}")

    workbook.close()
    laporan()
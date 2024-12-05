import openpyxl
import os

FILE_NAME = "Data_Nilai.xlsx"

def simpan_ke_excel(Nama_karyawan, Nama_Kegiatan, Nilai):
    # Jika file Excel tidak ada
    if not os.path.exists(FILE_NAME):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'Nilai'
        sheet.append(['Nama_karyawan', 'Nama_Kegiatan', 'Nilai'])
    else:
        # Jika sudah ada, buka file Excel
        workbook = openpyxl.load_workbook(FILE_NAME)
        sheet = workbook['Nilai']

    # Tambahkan data ke Excel
    sheet.append([Nama_karyawan, Nama_Kegiatan, Nilai])

    # Simpan perubahan
    workbook.save(FILE_NAME)

def niali():
    print("\nMAU MELAKUKAN APA DIniali???")
    print("1. Tambah Nilai")
    print("2. Lihat Nilai")
    print("3. Edit Nilai")
    print("4. Hapus Nilai")
    print("5. Kembali")
    pilihan = input("Pilih Menu: ")
    
    if pilihan == "1":
        tambah()
    elif pilihan == "2":
        baca()  
    elif pilihan == "3":
        edit()
    elif pilihan == "4":
        hapus()
    elif pilihan == "5":
        from main import menu_utama
        menu_utama()
    else:
        print("PILIHAN SALAH!!😒😒")
        
def tambah():
    print("\nIsi Untuk Menambahkan Nilai Karyawan!!")
    Nama_karyawan = input("Masukkan Nama Karyawan: ")
    Nama_Kegiatan = input("Masukkan Nama Kegiatan: ")
    Nilai = input("Masukkan Nilai: ")
    try:
        simpan_ke_excel(Nama_karyawan, Nama_Kegiatan, Nilai)
        print("Data berhasil ditambahkan ke Table Data Nilai.")
        niali()
    except Exception as e:
        print(f"Terjadi kesalahan saat menyimpan data: {e}")

def baca():
    if not os.path.exists(FILE_NAME):
        print("DATA TIDAK DITEMUKAN")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)
    if "Nilai" not in workbook.sheetnames:  # Check for the correct sheet name
        print("Sheet Belum ada")
        return
    
    sheet = workbook['Nilai']  # Use the correct sheet name
    if sheet.max_row == 1:
        print("Belum ada data")
        return
    
    print("\nDAFTAR KEGIATAN/Nilai")
    for row in sheet.iter_rows(min_row=2, values_only=True):
        print(f"Nama Karyawan: {row[0]}, \nNama Kegiatan: {row[1]}, \nNilai: {row[2]}\n")
    workbook.close()
    niali()

def edit():
    if not os.path.exists(FILE_NAME):
        print("DATA TIDAK DITEMUKAN")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)
    if "Nilai" not in workbook.sheetnames:
        print("Sheet Belum ada")
        return

    sheet = workbook['Nilai']
    if sheet.max_row == 1:
        print("Belum ada data untuk diedit.")
        return

    # Display current entries
    print("\nDAFTAR KEGIATAN/Nilai UNTUK DIEDIT")
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
        print(f"{index}. Nama Karyawan: {row[0]}, Nama Kegiatan: {row[1]}, Nilai: {row[2]}")

    # Ask user for the entry they want to edit
    try:
        pilihan = int(input("Masukkan nomor Nilai yang ingin diedit: "))
        if pilihan < 1 or pilihan > sheet.max_row - 1:  # Adjust for header row
            print("Nomor Nilai tidak valid.")
            return

        # Get the current data
        current_row = sheet[pilihan + 1]  # +1 because we skip the header row

        # Display current data
        print(f"Data saat ini: Nama Karyawan: {current_row[0].value}, Nama Kegiatan: {current_row[1].value}, Nilai: {current_row[2].value}")

        # Get new data from the user
        new_nama_karyawan = input("Masukkan Nama Karyawan baru (tekan Enter untuk tetap): ")
        new_nama_kegiatan = input("Masukkan Nama Kegiatan baru (tekan Enter untuk tetap): ")
        new_Nilai = input("Masukkan Nilai baru : ")

        # Update the row with new values
        if new_nama_karyawan:
            current_row[0].value = new_nama_karyawan
        if new_nama_kegiatan:
            current_row[1].value = new_nama_kegiatan
        if new_Nilai:
            current_row[2].value = new_Nilai

        # Save the changes
        workbook.save(FILE_NAME)
        print("Data berhasil diperbarui.")
    except Exception as e:
        print(f"Terjadi kesalahan saat mengedit data: {e}")
    finally:
        workbook.close()
    niali()
        
def hapus():
    if not os.path.exists(FILE_NAME):
        print("DATA TIDAK DITEMUKAN")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)
    if "Nilai" not in workbook.sheetnames:
        print("Sheet Belum ada")
        workbook.close()
        return

    sheet = workbook['Nilai']
    if sheet.max_row == 1:
        print("Belum ada data untuk dihapus.")
        workbook.close()
        return

    # Tampilkan data untuk membantu pengguna memilih
    print("\nDAFTAR KEGIATAN/NILAI UNTUK DIHAPUS")
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
        print(f"{index}. Nama Karyawan: {row[0]}, Nama Kegiatan: {row[1]}, Nilai: {row[2]}")

    try:
        # Meminta pengguna untuk memilih data yang ingin dihapus
        pilihan = int(input("Masukkan nomor data yang ingin dihapus: "))
        if pilihan < 1 or pilihan > sheet.max_row - 1:  # Validasi nomor urut
            print("Nomor tidak valid.")
            return

        # Hapus baris dari sheet
        sheet.delete_rows(pilihan + 1)  # +1 karena header ada di baris pertama

        # Simpan perubahan
        workbook.save(FILE_NAME)
        print("Data berhasil dihapus.")
    except ValueError:
        print("Input harus berupa angka.")
    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data: {e}")
    finally:
        workbook.close()
    niali()

import openpyxl
import os

from main import menu_utama

FILE_NAME = "Menu_Peserta.xlsx"
 
def peserta():
    menu_utama
    print("1. tambah peserta: ")
    print("2. Tampilkan Peserta: ")
    print("3. edit peserta: ")
    print("4. hapus peserta: ")
    print("5. kembali ke menu utama: ")
    pilih = input("mau yang mana??")

    if pilih =="1":
        tambah()
    elif pilih == "2":
        baca()
    elif pilih == "3":
        edit()
    elif pilih == "4":
        hapus()
    elif pilih ==  "5":
       menu_utama()
    else :
        print("salah")

def tambah():
    print("\nISI UNTUK MENAMBAHKAN NAMA PESERTA")
    nama_karyawan = input("Masukan Nama Karyawan: ")
    Umur = input("Masukan Umur Karyawan: ")
    telp = input("Masukan telp Karyawan: ")
    try:
        simpan_ke_excel(Umur, nama_karyawan, telp)
        print("Data berhasil ditambahkan ke menu peserta")
        peserta()
    except Exception as a:
        print(f"Terjadi kesalahan saat menyimpan data: {a}")


def baca():
    if not os.path.exists(FILE_NAME):
        print("DATA TIDAK DITEMUKAN")
        return
    
    workbook=openpyxl.load_workbook(FILE_NAME)
    if"Peserta" not in workbook.sheetnames:#check for the correct sheet name
        print("Sheet belum ada")
        return
    
    sheet=workbook['Peserta']#use the correct sheet name
    if sheet.max_row == 1:
        print("Belum ada data")
        return
    
    print("\nDAFTAR KEGIATAN/Peserta")
    for row in sheet.iter_rows(min_row=2, values_only=True):
        print(f"\nUmur:{row[0]}, \nNama Karyawan:{row[1]}, \ntelp:{row[2]}\n")
        workbook.close()
    peserta()
    

def edit():
    if not os.path.exists(FILE_NAME):
        print("DATA TIDAK DITEMUKAN")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)
    if "Peserta" not in workbook.sheetnames:
        print("Sheet Belum ada")
        return

    sheet = workbook['Peserta']
    if sheet.max_row == 1:
        print("Belum ada data untuk diedit.")
        return

    # Display current entries
    print("\nDAFTAR KEGIATAN UNTUK DIEDIT")
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
        print(f"{index}. Umur: {row[0]}, Nama Karyawan: {row[1]}, telp: {row[2]}")

    # Ask user for the entry they want to edit
    try:
        pilihan = int(input("Masukkan nomor peserta yang ingin diedit: "))
        if pilihan < 1 or pilihan > sheet.max_row - 1:  # Adjust for header row
            print("Nomor peserta tidak valid.")
            return

        # Get the current data
        current_row = sheet[pilihan + 1]  # +1 because we skip the header row

        # Display current data
        print(f"Data saat ini: Umur: {current_row[0].value}, Nama Karyawan: {current_row[1].value}, telp: {current_row[2].value}")

        # Get new data from the user
        new_Umur = input("Masukkan Umur baru (tekan Enter untuk tetap): ")
        new_nama_karyawan = input("Masukkan Nama Karyawan baru (tekan Enter untuk tetap): ")
        new_telp = input("Masukkan telp baru : ")

        # Update the row with new values
        if new_Umur:
            current_row[0].value = new_Umur
        if new_nama_karyawan:
            current_row[1].value = new_nama_karyawan
        if new_telp:
            current_row[2].value = new_telp

        # Save the changes
        workbook.save(FILE_NAME)
        print("Data berhasil diperbarui.")
    except Exception as e:
        print(f"Terjadi kesalahan saat mengedit data: {e}")
    finally:
        peserta()
        workbook.close()
   
                                

def hapus():
    if not os.path.exists(FILE_NAME):
        print("DATA TIDAK DITEMUKAN")
        return

    workbook = openpyxl.load_workbook(FILE_NAME)
    if "Peserta" not in workbook.sheetnames:
        print("Sheet Belum ada")
        return

    sheet = workbook['Peserta']
    if sheet.max_row == 1:
        print("Belum ada data untuk dihapus.")
        return

    # Display current entries
    print("\nDAFTAR KEGIATAN UNTUK DIHAPUS")
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
        print(f"{index}. Umur: {row[0]}, Nama karyawan: {row[1]}, telp: {row[2]}")

    # Ask user for the entry they want to delete
    try:
        pilihan = int(input("Masukkan nomor peserta yang ingin dihapus: "))
        if pilihan < 1 or pilihan > sheet.max_row - 1:  # Adjust for header row
            print("Nomor peserta tidak valid.")
            return

        # Delete the selected row
        sheet.delete_rows(pilihan + 1)  # +1 because we skip the header row

        # Save the changes
        workbook.save(FILE_NAME)
        print("Data berhasil dihapus.")
    except Exception as e:
        print(f"Terjadi kesalahan saat menghapus data: {e}")
    finally:
        peserta()
        workbook.close()


def simpan_ke_excel(Umur, Nama_Karyawan, telp):
    """
    Menyimpan data pembelian ke file Excel.
    """
    
    # Cek jika file Excel belum ada, buat file baru
    if not os.path.exists(FILE_NAME):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Peserta"
        
        # Tambahkan header
        sheet.append(["Umur", "Nama_Karyawan", "Jawaban"])
    else:
        workbook = openpyxl.load_workbook(FILE_NAME)
        sheet = workbook["Peserta"]
        
    # Tambahkan data pembelian
    sheet.append([Umur, Nama_Karyawan, telp])
    workbook.save(FILE_NAME)

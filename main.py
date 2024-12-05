def menu_utama():
    while True:
        print ("\nSISTEM PENGELOLAAN PELATIHAN KARYAWAN:") 
        print("1. Jadwal")
        print("2. Peserta")
        print("3. Nilai")
        print("4. Laporan Pelatihan")
        print("5. Keluar")
        pilihan = input ("Pilih menu: ")
        
        if pilihan == "1":
            from Menu_Jadwal import jadwal
            jadwal()
        elif pilihan == "2":
            from Manu_Peserta import peserta
            peserta()
        elif pilihan == "3":
            from Menu_Nilai import niali
            niali()
        elif pilihan == "4":
            from Menu_Laporan import laporan
            laporan()
        else:
            print("TERIMA KASIH SUDAH BERPARTISIPASI 🙏🙏AND BAYEEE❤️❤️")
        break
    
   
if __name__ == "__main__":
    menu_utama()

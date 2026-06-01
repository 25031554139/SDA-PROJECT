
# Kelompok 11
#Hariawan Napu(139)
#Aldian Dimas Ferdiansyah(224)
#Nandana Rizky Andromeda(252)

#IMPLEMENTASI STRUKTUR DATA QUEUE PADA SISTEM ANTRIAN BERBASIS PYTHON UNTUK SIMULASI PELAYANAN PUSKESMAS

import pandas as pd
from collections import deque
import os
import time

# Membaca file Excel yang berisi data antrian pasien
file_excel = "data_antrian.xlsx"

try:
    # Menggunakan pandas untuk membaca berkas Excel
    df = pd.read_excel(file_excel)
    # Mengonversi baris tabel Excel menjadi list of dictionaries agar mudah diolah dalam Python
    data_antrian_dummy = df.to_dict(orient="records")
    print(f"✅ Berhasil memuat {len(data_antrian_dummy)} data pasien dari Excel.\n")
except FileNotFoundError:
    # Antisipasi jika file Excel tidak ditemukan agar program tidak crash (tetap berjalan dengan data kosong)
    print(f"❌ File '{file_excel}' tidak ditemukan. Pastikan file ada di direktori yang sama.")
    data_antrian_dummy = []


# Fungsi pembantu untuk membuat visualisasi garis pembatas di terminal
def garis(karakter="═", panjang=55):
    print(karakter * panjang)


# Fungsi pembantu untuk mencetak judul menu di posisi tengah (center)
def header(judul):
    garis()
    print(f"  {judul.center(51)}")
    garis()


# Fungsi pembantu untuk menahan tampilan terminal sebelum dibersihkan kembali
def jeda():
    input("\n  [Tekan ENTER untuk lanjut...]")
    # Membersihkan layar terminal (cls untuk Windows, clear untuk Linux/Mac)
    os.system("cls" if os.name == "nt" else "clear")


# Kelas utama untuk mengelola seluruh logika sistem antrean puskesmas
class SistemAntrianPuskesmas:
    BATAS_PANGGILAN = 3 # Konstanta batas maksimal pemanggilan pasien yang absen
    NAMA_POLI = {
        "1": "Poliklinik Umum",
        "2": "Poliklinik Gigi"
    }

    # Fungsi inisialisasi objek (Constructor) untuk menyiapkan struktur data awal
    def __init__(self, jumlah_loket_umum=3, jumlah_loket_gigi=2):
        # Menggunakan Dictionary yang berisi List dari objek deque() untuk menerapkan struktur data Queue (FIFO)
        # Poliklinik Umum memiliki 3 loket mandiri, Poliklinik Gigi memiliki 2 loket mandiri
        self.antrean = {
            "Poliklinik Umum": [deque() for _ in range(jumlah_loket_umum)],
            "Poliklinik Gigi": [deque() for _ in range(jumlah_loket_gigi)]
        }
        self.pasien_terlewat = [] # List untuk menampung data pasien yang batal/hangus
        # Cache/Memori jangka pendek untuk menyimpan pasien yang baru saja dipanggil berdasarkan (poli, loket)
        self.pasien_terakhir_dipanggil = {}
        self.counter_nomor_antrian = 1000 # Base counter untuk pembuatan nomor antrean pasien manual

    # ------------------------------------------------------------
    # DAFTAR PASIEN (dari Excel / manual)
    # LOGIKA LOKET PALING SEDIKIT (LOAD BALANCING)
    # ------------------------------------------------------------
    def tambah_ke_antrean(self, data_pasien):
        # Set nilai awal jumlah_panggilan = 0 jika belum ada di dalam data pasien
        data_pasien.setdefault("jumlah_panggilan", 0)
        poli_tujuan = data_pasien.get("poli")
        
        # Validasi apakah poliklinik tujuan terdaftar dalam sistem
        if poli_tujuan in self.antrean:
            daftar_loket = self.antrean[poli_tujuan]
            
            # LOGIKA UTAMA LOKET PALING SEDIKIT:
            # Fungsi min() akan mencari indeks loket (i) yang memiliki panjang antrean (len) paling pendek/sedikit.
            loket_idx = min(range(len(daftar_loket)), key=lambda i: len(daftar_loket[i]))
            
            # Memasukkan data pasien ke ujung belakang (append) pada loket yang paling sepi tersebut (Aspek FIFO)
            self.antrean[poli_tujuan][loket_idx].append(data_pasien)
            print(f"  ✔  {data_pasien['nama']:<25} → {poli_tujuan} | Loket {loket_idx + 1} | No. {data_pasien['nomor_antrian']}")
        else:
            print(f"  ✘  Poli tidak valid untuk pasien: {data_pasien['nama']}")

    # Fungsi Menu [5]: Mengelola input data pasien baru yang datang langsung (walk-in)
    def daftar_pasien_manual(self):
        print("\n  Masukkan data pasien baru:")
        print("  (Ketik '0' di field nama untuk membatalkan)\n")
        nama = input("  Nama Pasien       : ").strip()
        if nama == "0":
            return None
        if not nama:
            print("  ⚠  Nama tidak boleh kosong.")
            return None

        # Memanggil fungsi internal untuk memilih poliklinik tujuan
        poli = self.pilih_poli()
        if not poli:
            return None

        print("\n  Pilih Keperluan:")
        daftar_keperluan = {
            "1": "Pemeriksaan Umum",
            "2": "Kontrol Rutin",
            "3": "Pengambilan Obat",
            "4": "Lainnya"
        }
        for kode, nama_keperluan in daftar_keperluan.items():
            print(f"    [{kode}] {nama_keperluan}")

        keperluan = "Lainnya"
        while True:
            pilihan_keperluan = input("\n  Masukkan pilihan keperluan: ").strip()
            if pilihan_keperluan in daftar_keperluan:
                keperluan = daftar_keperluan[pilihan_keperluan]
                break
            print("  ⚠  Pilihan tidak valid. Coba lagi.")

        keterangan = input("\n  Keterangan Tambahan (opsional): ").strip()
        if not keterangan:
            keterangan = "-"

        # Membuat format kode antrean manual "M-XXXX" secara berurutan
        nomor_antrian = f"M-{self.counter_nomor_antrian}"
        self.counter_nomor_antrian += 1

        # Menyusun data ke dalam struktur dictionary pasien standar
        data_pasien = {
            "nama": nama,
            "nomor_antrian": nomor_antrian,
            "poli": poli,
            "keperluan": keperluan,
            "keterangan_keperluan": keterangan,
            "jumlah_panggilan": 0
        }

        garis("─")
        print(f"\n  ✅  Pasien berhasil didaftarkan!")
        print(f"  Nama          : {nama}")
        print(f"  No. Antrian   : {nomor_antrian}")
        print(f"  Poliklinik    : {poli}")
        print(f"  Keperluan     : {keperluan}")
        garis("─")

        # Memasukkan pasien baru tersebut ke loket paling sedikit melalui fungsi yang sudah dibuat
        self.tambah_ke_antrean(data_pasien)
        return data_pasien

    # ------------------------------------------------------------
    # PILIH POLI & LOKET (Fungsi pembantu menu operasional)
    # ------------------------------------------------------------
    def pilih_poli(self):
        print("\n  Pilih Poliklinik:")
        for kode, nama in self.NAMA_POLI.items():
            jumlah_loket = len(self.antrean[nama])
            print(f"    [{kode}] {nama} ({jumlah_loket} loket)")
        print("    [0] Kembali")
        while True:
            pilihan = input("\n  Masukkan pilihan: ").strip()
            if pilihan == "0":
                return None
            if pilihan in self.NAMA_POLI:
                return self.NAMA_POLI[pilihan]
            print("  ⚠  Pilihan tidak valid. Coba lagi.")

    def pilih_loket(self, jenis_poli):
        daftar_loket = self.antrean[jenis_poli]
        print(f"\n  Pilih Loket — {jenis_poli}:")
        for i, q in enumerate(daftar_loket):
            status = f"{len(q)} antrian"
            print(f"    [{i + 1}] Loket {i + 1}  ({status})")
        print("    [0] Kembali")
        while True:
            try:
                pilihan = input("\n  Masukkan nomor loket: ").strip()
                if pilihan == "0":
                    return None
                nomor = int(pilihan)
                if 1 <= nomor <= len(daftar_loket):
                    return nomor
                print(f"  ⚠  Loket harus antara 1–{len(daftar_loket)}.")
            except ValueError:
                print("  ⚠  Masukkan angka yang valid.")

    # ------------------------------------------------------------
    # PANGGIL PASIEN (OPERASI DEQUEUE)
    # ------------------------------------------------------------
    def panggil_pasien(self, jenis_poli, nomor_loket):
        # Mengakses objek antrean loket yang dituju
        queue_loket = self.antrean[jenis_poli][nomor_loket - 1]
        
        # Validasi jika antrean pada loket tersebut kosong
        if not queue_loket:
            print(f"\n  ℹ  Antrean di {jenis_poli} Loket {nomor_loket} sudah KOSONG.")
            return None

        # OPERASI DEQUEUE UTAMA: Mengambil dan menghapus pasien di urutan PALING DEPAN (popleft) sesuai asas FIFO
        pasien = queue_loket.popleft()
        pasien["jumlah_panggilan"] += 1 # Menambahkan riwayat pemanggilan pasien

        # Menyimpan data pasien ke dalam cache sementara (pasien_terakhir_dipanggil) dengan status awal None.
        # Ini berfungsi agar data pasien tidak hilang dari sistem sebelum diverifikasi kehadirannya.
        kunci = (jenis_poli, nomor_loket)
        self.pasien_terakhir_dipanggil[kunci] = {
            'data': pasien,
            'status': None
        }

        # Menampilkan detail data pasien ke layar monitor petugas
        garis("─")
        print(f"  🔔  PANGGILAN KE-{pasien['jumlah_panggilan']}")
        garis("─")
        print(f"  Nama          : {pasien['nama']}")
        print(f"  No. Antrian   : {pasien['nomor_antrian']}")
        print(f"  Keperluan     : {pasien['keperluan']}")
        print(f"  Keterangan    : {pasien['keterangan_keperluan']}")
        print(f"  Sisa antrian  : {len(queue_loket)} orang di loket ini")
        garis("─")

        return pasien

    # ------------------------------------------------------------
    # KONFIRMASI HADIR (Menandai status di cache)
    # ------------------------------------------------------------
    def tandai_hadir(self, jenis_poli, nomor_loket):
        kunci = (jenis_poli, nomor_loket)
        # Mengubah status pasien di cache menjadi 'hadir' jika pasien datang ke loket
        if kunci in self.pasien_terakhir_dipanggil:
            self.pasien_terakhir_dipanggil[kunci]['status'] = 'hadir'

    # ------------------------------------------------------------
    # CODE KETIKA PASIEN TIDAK ADA SAAT DIPANGGIL (BATAS 3 KALI)
    # ------------------------------------------------------------
    def tangani_tidak_hadir(self, pasien, jenis_poli, nomor_loket):
        if not pasien:
            return
        sisa = self.BATAS_PANGGILAN - pasien["jumlah_panggilan"]
        print(f"\n  ⚠  {pasien['nama']} (No. {pasien['nomor_antrian']}) TIDAK HADIR")
        print(f"     Panggilan ke-{pasien['jumlah_panggilan']} dari maks {self.BATAS_PANGGILAN}x")

        # KONDISI JIKA SUDAH DIPANGGIL 3 KALI ATAU LEBIH:
        if pasien["jumlah_panggilan"] >= self.BATAS_PANGGILAN:
            # Nomor antrean dinyatakan hangus/batal dan dipindahkan permanen ke list pasien_terlewat
            self.pasien_terlewat.append(pasien)
            print(f"  ✘  Batas tercapai. Pasien masuk DAFTAR TERLEWAT.")
        # KONDISI JIKA BELUM MENCAPAI 3 KALI PANGGILAN:
        else:
            # Sesuai asas keadilan antrean, pasien dimasukkan kembali ke urutan PALING BELAKANG (.append) pada loketnya semula
            self.antrean[jenis_poli][nomor_loket - 1].append(pasien)
            print(f"  ↩  Pasien dikembalikan ke belakang antrian. Sisa kesempatan: {sisa}x.")

    # ------------------------------------------------------------
    # KOREKSI DARI MENU 2: Mengambil dari cache dan diproses tidak hadir
    # ------------------------------------------------------------
    def koreksi_tidak_hadir_dari_cache(self, jenis_poli, nomor_loket):
        kunci = (jenis_poli, nomor_loket)
        item = self.pasien_terakhir_dipanggil.get(kunci)
        if not item:
            return False
        pasien = item['data']
        # Mengarahkan pasien ke fungsi penanganan tidak hadir (ditunda atau dibatalkan)
        self.tangani_tidak_hadir(pasien, jenis_poli, nomor_loket)
        # Menghapus entri dari cache karena penanganan koreksi salah input sudah selesai dilakukan
        del self.pasien_terakhir_dipanggil[kunci]
        return True

    # Menghapus data dari cache jika pasien dikonfirmasi tidak hadir langsung dari Menu 1
    def hapus_dari_cache(self, jenis_poli, nomor_loket):
        kunci = (jenis_poli, nomor_loket)
        if kunci in self.pasien_terakhir_dipanggil:
            del self.pasien_terakhir_dipanggil[kunci]

    # ------------------------------------------------------------
    # Menu [3]: LIHAT STATUS ANTRIAN (Monitoring Sistem)
    # ------------------------------------------------------------
    def tampilkan_status(self):
        header("📋  STATUS ANTRIAN SAAT INI")
        # Melakukan looping untuk menampilkan kondisi grafis antrean aktif di setiap poliklinik dan loket
        for poli, daftar_loket in self.antrean.items():
            print(f"\n  {poli}")
            garis("─", 40)
            total = 0
            for i, q in enumerate(daftar_loket):
                total += len(q)
                # Membuat visualisasi bar grafik sederhana menggunakan karakter balok "█"
                bar = "█" * len(q) if q else "-"
                print(f"    Loket {i + 1}  : {len(q):>3} orang  {bar}")
            print(f"    {'TOTAL':<9}: {total:>3} orang")

        # Menampilkan ringkasan data daftar pasien yang hangus/batal (terlewat 3x)
        garis("─", 40)
        print(f"\n  Pasien Terlewat / Batal : {len(self.pasien_terlewat)} orang")
        if self.pasien_terlewat:
            print()
            for p in self.pasien_terlewat:
                print(f"    ✘ {p['nama']:<25} No. {p['nomor_antrian']}  (dipanggil {p['jumlah_panggilan']}x)")

        # Menampilkan log pasien yang berada di dalam cache monitoring pengawasan petugas
        if self.pasien_terakhir_dipanggil:
            print(f"\n  Pasien yang sudah dipanggil dan belum selesai (cache):")
            garis("─", 40)
            for (poli, loket), item in self.pasien_terakhir_dipanggil.items():
                p = item['data']
                status_text = "HADIR" if item['status'] == 'hadir' else "BELUM dikonfirmasi"
                print(f"    {p['nama']:<25} → {poli} Loket {loket}  [{status_text}]")
        garis()


# ══════════════════════════════════════════════════════
# PENJELASAN ALUR PADA MASING-MASING MENU UTAMA
# ══════════════════════════════════════════════════════
def menu_utama(puskesmas):
    while True:
        header("🏥  SISTEM ANTRIAN PUSKESMAS")
        print("  [1] Panggil Pasien")
        print("  [2] Koreksi Tidak Hadir (untuk pasien yang masih dalam cache)")
        print("  [3] Lihat Status Antrian")
        print("  [4] Daftar Ulang Pasien dari Excel")
        print("  [5] Daftarkan Pasien Baru")
        print("  [0] Keluar")
        garis()

        pilihan = input("  Pilih menu: ").strip()

        # ---------- PENJELASAN MENU 1 : PANGGIL PASIEN ----------
        if pilihan == "1":
            os.system("cls" if os.name == "nt" else "clear")
            header("🔔  PANGGIL PASIEN")
            poli = puskesmas.pilih_poli()
            if not poli:
                continue
            loket = puskesmas.pilih_loket(poli)
            if not loket:
                continue
            print()
            pasien = puskesmas.panggil_pasien(poli, loket)

            # Setelah memanggil pasien terdepan, operator wajib mengonfirmasi kehadiran fisik pasien tersebut
            if pasien:
                print("\n  Apakah pasien HADIR?")
                print("  [1] Ya, hadir")
                print("  [2] Tidak hadir")
                konfirmasi = input("\n  Pilihan: ").strip()
                if konfirmasi == "2":
                    # Pilihan [2]: Pasien absen -> langsung diproses mundur/batal dan datanya dihapus dari cache aktif
                    puskesmas.tangani_tidak_hadir(pasien, poli, loket)
                    puskesmas.hapus_dari_cache(poli, loket)
                else:
                    # Pilihan [1]: Pasien ada -> status cache diset 'hadir' (masih disimpan di cache untuk cadangan Menu 2)
                    puskesmas.tandai_hadir(poli, loket)
                    print(f"\n  ✅  {pasien['nama']} tercatat HADIR.")
                    print("  ℹ️  Jika keliru, segera gunakan Menu [2] untuk mengubah menjadi TIDAK HADIR.")
            jeda()

        # ---------- PENJELASAN MENU 2 : KOREKSI TIDAK HADIR (Penyelamat salah input) ----------
        elif pilihan == "2":
            os.system("cls" if os.name == "nt" else "clear")
            header("⚠️  KOREKSI TIDAK HADIR")

            # Memvalidasi ketersediaan log cache data. Menu ini hanya bekerja jika cache tidak kosong.
            if not puskesmas.pasien_terakhir_dipanggil:
                print("\n  ℹ  Tidak ada pasien yang perlu dikoreksi saat ini.")
                print("     (Cache kosong)")
                jeda()
                continue

            # Menampilkan daftar pasien terakhir yang baru dikonfirmasi 'Hadir' dari Menu 1
            print("\n  Pasien yang sudah dipanggil dan belum dihapus dari cache:\n")
            daftar_kunci = list(puskesmas.pasien_terakhir_dipanggil.keys())
            for idx, (poli, loket) in enumerate(daftar_kunci, start=1):
                item = puskesmas.pasien_terakhir_dipanggil[(poli, loket)]
                p = item['data']
                status_text = "HADIR" if item['status'] == 'hadir' else "BELUM dikonfirmasi"
                print(f"  [{idx}] {p['nama']:<25} → {poli} Loket {loket}  (No. {p['nomor_antrian']})  [{status_text}]")
            print("  [0] Kembali")

            try:
                pilih_koreksi = input("\n  Pilih nomor pasien yang ternyata TIDAK HADIR: ").strip()
                if pilih_koreksi == "0":
                    continue
                nomor = int(pilih_koreksi)
                if 1 <= nomor <= len(daftar_kunci):
                    poli_terpilih, loket_terpilih = daftar_kunci[nomor - 1]
                    print()
                    # Menjalankan fungsi koreksi untuk mengubah status pasien menjadi Tidak Hadir (mundur ke belakang/batal)
                    berhasil = puskesmas.koreksi_tidak_hadir_dari_cache(poli_terpilih, loket_terpilih)
                    if berhasil:
                        print(f"\n  ✅  Pasien berhasil diproses sebagai TIDAK HADIR.")
                    else:
                        print("\n  ✘  Gagal memproses. Mungkin cache sudah tidak ada.")
                else:
                    print(f"  ⚠  Pilihan harus antara 1–{len(daftar_kunci)}.")
            except ValueError:
                print("  ⚠  Masukkan angka yang valid.")
            jeda()

        # ---------- PENJELASAN MENU 3 : LIHAT STATUS ----------
        elif pilihan == "3":
            os.system("cls" if os.name == "nt" else "clear")
            puskesmas.tampilkan_status() # Memanggil fungsi internal monitoring grafis antrean
            jeda()

        # ---------- PENJELASAN MENU 4 : RELOAD EXCEL (Daftar Ulang) ----------
        elif pilihan == "4":
            os.system("cls" if os.name == "nt" else "clear")
            header("📂  DAFTAR ULANG DARI EXCEL")
            # Memasukkan kembali seluruh baris data pendaftaran awal Excel ke dalam sistem pembagian loket
            if data_antrian_dummy:
                for pasien in data_antrian_dummy:
                    puskesmas.tambah_ke_antrean(pasien)
                print("\n  ✅ Semua pasien dari Excel telah didaftarkan ulang.")
            else:
                print("  ❌ Tidak ada data Excel yang tersedia.")
            jeda()

        # ---------- PENJELASAN MENU 5 : DAFTAR PASIEN MANUAL ----------
        elif pilihan == "5":
            os.system("cls" if os.name == "nt" else "clear")
            header("📝  DAFTARKAN PASIEN BARU")
            puskesmas.daftar_pasien_manual() # Memanggil runtutan input form pendaftaran manual walk-in
            jeda()

        # ---------- KELUAR DARI PROGRAM ----------
        elif pilihan == "0":
            print("\n  Sampai jumpa! 👋\n")
            break

        else:
            print("  ⚠  Pilihan tidak valid. Coba lagi.")
            time.sleep(1)
            os.system("cls" if os.name == "nt" else "clear")


# ══════════════════════════════════════════════════════
# ENTRY POINT (Gerbang Utama Eksekusi Program)
# ══════════════════════════════════════════════════════
if __name__ == "__main__":
    os.system("cls" if os.name == "nt" else "clear")
    # Instansiasi objek struktur data sistem antrean puskesmas
    puskesmas = SistemAntrianPuskesmas(jumlah_loket_umum=3, jumlah_loket_gigi=2)

    # Memproses otomasi pendaftaran awal jika data dari file Excel berhasil dimuat saat startup program
    if data_antrian_dummy:
        header("📂  PENDAFTARAN AWAL PASIEN")
        for pasien in data_antrian_dummy:
            puskesmas.tambah_ke_antrean(pasien)
        jeda()

    # Membuka gerbang antarmuka perulangan menu utama untuk operator puskesmas
    menu_utama(puskesmas)

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
    df = pd.read_excel(file_excel)
    # Mengubah DataFrame menjadi list of dictionary agar mudah diproses
    data_antrian_dummy = df.to_dict(orient="records")
    print(f"✅ Berhasil memuat {len(data_antrian_dummy)} data pasien dari Excel.\n")
except FileNotFoundError:
    print(f"❌ File '{file_excel}' tidak ditemukan. Pastikan file ada di direktori yang sama.")
    data_antrian_dummy = []


def garis(karakter="═", panjang=55):
    """Mencetak garis pemisah dengan karakter dan panjang tertentu."""
    print(karakter * panjang)


def header(judul):
    """Mencetak header dengan judul yang di-tengahkan, diapit garis."""
    garis()
    print(f"  {judul.center(51)}")
    garis()


def jeda():
    """Menunggu input ENTER dari pengguna, lalu membersihkan layar."""
    input("\n  [Tekan ENTER untuk lanjut...]")
    os.system("cls" if os.name == "nt" else "clear")


class SistemAntrianPuskesmas:
    """
    Sistem manajemen antrian untuk puskesmas dengan beberapa poliklinik dan loket.

    Sistem ini mendukung:
    - Distribusi pasien otomatis ke loket yang paling sedikit antriannya.
    - Pemanggilan pasien per loket.
    - Penanganan pasien tidak hadir dengan batas maksimal pemanggilan.
    - Pencatatan pasien terlewat/batal.
    - Pendaftaran pasien baru secara manual melalui input konsol.

    Attributes:
        BATAS_PANGGILAN (int): Jumlah maksimal pemanggilan sebelum pasien dinyatakan terlewat.
        NAMA_POLI (dict): Mapping kode pilihan ke nama poliklinik.
        antrean (dict): Struktur antrian per poliklinik dan loket.
        pasien_terlewat (list): Daftar pasien yang telah melewati batas pemanggilan.
        pasien_terakhir_dipanggil (dict): Cache pasien terakhir yang dipanggil per (poli, loket).
        counter_nomor_antrian (int): Counter otomatis untuk nomor antrian pasien yang didaftarkan manual.
    """

    # Batas maksimal pasien boleh dipanggil sebelum dinyatakan batal
    BATAS_PANGGILAN = 3

    # Mapping kode input ke nama poliklinik yang tersedia
    NAMA_POLI = {
        "1": "Poliklinik Umum",
        "2": "Poliklinik Gigi"
    }

    def __init__(self, jumlah_loket_umum=3, jumlah_loket_gigi=2):
        """
        Inisialisasi sistem antrian puskesmas.

        Args:
            jumlah_loket_umum (int): Jumlah loket di Poliklinik Umum. Default 3.
            jumlah_loket_gigi (int): Jumlah loket di Poliklinik Gigi. Default 2.
        """
        # Setiap poli memiliki beberapa loket, masing-masing berupa deque (Queue)
        self.antrean = {
            "Poliklinik Umum": [deque() for _ in range(jumlah_loket_umum)],
            "Poliklinik Gigi":  [deque() for _ in range(jumlah_loket_gigi)]
        }
        # Menyimpan pasien yang sudah mencapai batas pemanggilan dan dinyatakan batal
        self.pasien_terlewat = []
        # Menyimpan pasien terakhir yang dipanggil per (poli, loket)
        # agar bisa ditandai tidak hadir dari menu terpisah
        self.pasien_terakhir_dipanggil = {}
        # Counter otomatis untuk nomor antrian pasien manual, dimulai dari M-1000
        self.counter_nomor_antrian = 1000

    # ──────────────────────────────────────────
    # DAFTAR PASIEN
    # ──────────────────────────────────────────
    def tambah_ke_antrean(self, data_pasien):
        """
        Mendaftarkan pasien ke loket dengan antrian paling sedikit (load balancing).

        Args:
            data_pasien (dict): Data pasien yang berisi 'nama', 'nomor_antrian',
                                'poli', 'keperluan', dan 'keterangan_keperluan'.
        """
        # Inisialisasi counter panggilan jika belum ada
        data_pasien.setdefault("jumlah_panggilan", 0)
        poli_tujuan = data_pasien.get("poli")

        if poli_tujuan in self.antrean:
            daftar_loket = self.antrean[poli_tujuan]
            # Cari loket dengan antrian terpendek menggunakan fungsi min()
            loket_idx = min(range(len(daftar_loket)), key=lambda i: len(daftar_loket[i]))
            # Masukkan pasien ke belakang antrian loket terpilih (enqueue)
            self.antrean[poli_tujuan][loket_idx].append(data_pasien)
            print(f"  ✔  {data_pasien['nama']:<25} → {poli_tujuan} | Loket {loket_idx + 1} | No. {data_pasien['nomor_antrian']}")
        else:
            print(f"  ✘  Poli tidak valid untuk pasien: {data_pasien['nama']}")

    # ──────────────────────────────────────────
    # DAFTAR PASIEN MANUAL
    # ──────────────────────────────────────────
    def daftar_pasien_manual(self):
        """
        Mendaftarkan pasien baru secara manual melalui input konsol.

        Operator diminta mengisi nama, poliklinik tujuan, keperluan, dan
        keterangan keperluan. Nomor antrian dibuat otomatis menggunakan
        counter internal yang dimulai dari M-1000 untuk membedakan dari
        pasien yang didaftarkan via Excel.

        Returns:
            dict | None: Data pasien yang berhasil didaftarkan,
                         atau None jika operator membatalkan pendaftaran.
        """
        print("\n  Masukkan data pasien baru:")
        print("  (Ketik '0' di field nama untuk membatalkan)\n")

        # Input nama pasien, '0' digunakan sebagai sinyal pembatalan
        nama = input("  Nama Pasien       : ").strip()
        if nama == "0":
            return None
        if not nama:
            print("  ⚠  Nama tidak boleh kosong.")
            return None

        # Pilih poliklinik tujuan melalui menu interaktif
        poli = self.pilih_poli()
        if not poli:
            return None

        # Tampilkan pilihan keperluan yang tersedia
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

        # Keterangan tambahan bersifat opsional, diisi '-' jika kosong
        keterangan = input("\n  Keterangan Tambahan (opsional): ").strip()
        if not keterangan:
            keterangan = "-"

        # Generate nomor antrian otomatis dengan prefix 'M' untuk pasien manual
        nomor_antrian = f"M-{self.counter_nomor_antrian}"
        self.counter_nomor_antrian += 1

        data_pasien = {
            "nama"                 : nama,
            "nomor_antrian"        : nomor_antrian,
            "poli"                 : poli,
            "keperluan"            : keperluan,
            "keterangan_keperluan" : keterangan,
            "jumlah_panggilan"     : 0
        }

        garis("─")
        print(f"\n  ✅  Pasien berhasil didaftarkan!")
        print(f"  Nama          : {nama}")
        print(f"  No. Antrian   : {nomor_antrian}")
        print(f"  Poliklinik    : {poli}")
        print(f"  Keperluan     : {keperluan}")
        garis("─")

        self.tambah_ke_antrean(data_pasien)
        return data_pasien

    # ──────────────────────────────────────────
    # PILIH POLI (interaktif)
    # ──────────────────────────────────────────
    def pilih_poli(self):
        """
        Menampilkan daftar poliklinik dan meminta pengguna memilih salah satu.

        Returns:
            str | None: Nama poliklinik yang dipilih, atau None jika pengguna membatalkan.
        """
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

    # ──────────────────────────────────────────
    # PILIH LOKET (interaktif)
    # ──────────────────────────────────────────
    def pilih_loket(self, jenis_poli):
        """
        Menampilkan daftar loket dari poliklinik tertentu dan meminta pengguna memilih.

        Args:
            jenis_poli (str): Nama poliklinik yang loketnya ingin dipilih.

        Returns:
            int | None: Nomor loket yang dipilih (1-based), atau None jika dibatalkan.
        """
        daftar_loket = self.antrean[jenis_poli]
        print(f"\n  Pilih Loket — {jenis_poli}:")
        for i, q in enumerate(daftar_loket):
            # Tampilkan jumlah antrian saat ini di setiap loket sebagai informasi operator
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

    # ──────────────────────────────────────────
    # PANGGIL PASIEN
    # ──────────────────────────────────────────
    def panggil_pasien(self, jenis_poli, nomor_loket):
        """
        Memanggil pasien berikutnya dari antrian loket yang ditentukan.

        Pasien yang dipanggil disimpan sementara di `pasien_terakhir_dipanggil`
        agar bisa ditandai tidak hadir melalui menu terpisah jika diperlukan.

        Args:
            jenis_poli (str): Nama poliklinik.
            nomor_loket (int): Nomor loket (1-based).

        Returns:
            dict | None: Data pasien yang dipanggil, atau None jika antrian kosong.
        """
        queue_loket = self.antrean[jenis_poli][nomor_loket - 1]

        if not queue_loket:
            print(f"\n  ℹ  Antrean di {jenis_poli} Loket {nomor_loket} sudah KOSONG.")
            return None

        # Ambil pasien dari depan antrian — operasi dequeue O(1)
        pasien = queue_loket.popleft()
        pasien["jumlah_panggilan"] += 1

        # Simpan ke cache agar bisa diakses dari Menu [2] jika perlu koreksi
        kunci = (jenis_poli, nomor_loket)
        self.pasien_terakhir_dipanggil[kunci] = pasien

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

    # ──────────────────────────────────────────
    # TANDAI HADIR (hapus dari cache)
    # ──────────────────────────────────────────
    def tandai_hadir(self, jenis_poli, nomor_loket):
        """
        Menandai pasien terakhir yang dipanggil sebagai hadir dan menghapusnya dari cache.

        Args:
            jenis_poli (str): Nama poliklinik.
            nomor_loket (int): Nomor loket (1-based).
        """
        kunci = (jenis_poli, nomor_loket)
        # Hapus dari cache karena pasien sudah terkonfirmasi hadir
        if kunci in self.pasien_terakhir_dipanggil:
            del self.pasien_terakhir_dipanggil[kunci]

    # ──────────────────────────────────────────
    # TANDAI TIDAK HADIR (dari cache)
    # ──────────────────────────────────────────
    def tandai_tidak_hadir_dari_cache(self, jenis_poli, nomor_loket):
        """
        Menandai pasien terakhir yang dipanggil (dari cache) sebagai tidak hadir.

        Digunakan oleh Menu [2] sebagai tindakan koreksi jika operator lupa
        menandai tidak hadir saat pemanggilan berlangsung.

        Args:
            jenis_poli (str): Nama poliklinik.
            nomor_loket (int): Nomor loket (1-based).

        Returns:
            bool: True jika berhasil, False jika tidak ada pasien di cache.
        """
        kunci = (jenis_poli, nomor_loket)
        pasien = self.pasien_terakhir_dipanggil.get(kunci)

        if not pasien:
            print(f"\n  ℹ  Tidak ada pasien yang menunggu konfirmasi di {jenis_poli} Loket {nomor_loket}.")
            print("     Pastikan pasien sudah dipanggil terlebih dahulu melalui Menu [1].")
            return False

        self.tangani_tidak_hadir(pasien, jenis_poli, nomor_loket)
        # Hapus dari cache setelah status tidak hadir selesai diproses
        del self.pasien_terakhir_dipanggil[kunci]
        return True

    # ──────────────────────────────────────────
    # TANGANI TIDAK HADIR
    # ──────────────────────────────────────────
    def tangani_tidak_hadir(self, pasien, jenis_poli, nomor_loket):
        """
        Memproses pasien yang tidak hadir setelah dipanggil.

        Jika jumlah pemanggilan belum mencapai BATAS_PANGGILAN, pasien dikembalikan
        ke belakang antrian. Jika sudah mencapai batas, pasien dimasukkan ke daftar terlewat.

        Args:
            pasien (dict): Data pasien yang tidak hadir.
            jenis_poli (str): Nama poliklinik.
            nomor_loket (int): Nomor loket (1-based).
        """
        if not pasien:
            return

        sisa = self.BATAS_PANGGILAN - pasien["jumlah_panggilan"]
        print(f"\n  ⚠  {pasien['nama']} (No. {pasien['nomor_antrian']}) TIDAK HADIR")
        print(f"     Panggilan ke-{pasien['jumlah_panggilan']} dari maks {self.BATAS_PANGGILAN}x")

        if pasien["jumlah_panggilan"] >= self.BATAS_PANGGILAN:
            # Batas tercapai — pasien dicoret dan tidak akan dipanggil lagi
            self.pasien_terlewat.append(pasien)
            print(f"  ✘  Batas tercapai. Pasien masuk DAFTAR TERLEWAT.")
        else:
            # Masih ada kesempatan — kembalikan pasien ke belakang antrian (re-enqueue)
            self.antrean[jenis_poli][nomor_loket - 1].append(pasien)
            print(f"  ↩  Pasien dikembalikan ke belakang antrian. Sisa kesempatan: {sisa}x.")

    # ──────────────────────────────────────────
    # STATUS ANTRIAN
    # ──────────────────────────────────────────
    def tampilkan_status(self):
        """
        Menampilkan ringkasan status antrian seluruh poliklinik dan loket,
        termasuk daftar pasien yang terlewat/batal.
        """
        header("📋  STATUS ANTRIAN SAAT INI")
        for poli, daftar_loket in self.antrean.items():
            print(f"\n  {poli}")
            garis("─", 40)
            total = 0
            for i, q in enumerate(daftar_loket):
                total += len(q)
                # Visualisasi sederhana jumlah antrian menggunakan karakter blok
                bar = "█" * len(q) if q else "-"
                print(f"    Loket {i + 1}  : {len(q):>3} orang  {bar}")
            print(f"    {'TOTAL':<9}: {total:>3} orang")

        garis("─", 40)
        print(f"\n  Pasien Terlewat / Batal : {len(self.pasien_terlewat)} orang")
        if self.pasien_terlewat:
            print()
            for p in self.pasien_terlewat:
                print(f"    ✘ {p['nama']:<25} No. {p['nomor_antrian']}  (dipanggil {p['jumlah_panggilan']}x)")

        # Tampilkan pasien yang masih menunggu konfirmasi kehadiran dari operator
        if self.pasien_terakhir_dipanggil:
            print(f"\n  Menunggu Konfirmasi Kehadiran:")
            garis("─", 40)
            for (poli, loket), p in self.pasien_terakhir_dipanggil.items():
                print(f"    ⏳ {p['nama']:<25} → {poli} Loket {loket}")

        garis()


# ══════════════════════════════════════════════════════
# MENU UTAMA
# ══════════════════════════════════════════════════════
def menu_utama(puskesmas):
    """
    Menjalankan loop menu utama sistem antrian puskesmas.

    Menu yang tersedia:
        [1] Panggil Pasien — memanggil pasien berikutnya dari loket yang dipilih.
        [2] Tandai Tidak Hadir — menandai pasien terakhir yang dipanggil sebagai tidak hadir.
        [3] Lihat Status Antrian — menampilkan ringkasan antrian semua loket.
        [4] Daftar Ulang Pasien dari Excel — memuat ulang semua data dari file Excel.
        [5] Daftarkan Pasien Baru — mendaftarkan pasien baru secara manual via input konsol.
        [0] Keluar — menutup aplikasi.

    Args:
        puskesmas (SistemAntrianPuskesmas): Objek sistem antrian yang aktif.
    """
    while True:
        header("🏥  SISTEM ANTRIAN PUSKESMAS")
        print("  [1] Panggil Pasien")
        print("  [2] Tandai Tidak Hadir  (koreksi pasien yang baru dipanggil)")
        print("  [3] Lihat Status Antrian")
        print("  [4] Daftar Ulang Pasien dari Excel")
        print("  [5] Daftarkan Pasien Baru")
        print("  [0] Keluar")
        garis()

        pilihan = input("  Pilih menu: ").strip()

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

            if pasien:
                print("\n  Apakah pasien HADIR?")
                print("  [1] Ya, hadir")
                print("  [2] Tidak hadir")
                konfirmasi = input("\n  Pilihan: ").strip()
                if konfirmasi == "2":
                    puskesmas.tangani_tidak_hadir(pasien, poli, loket)
                    # Hapus dari cache karena sudah diproses langsung di menu ini
                    puskesmas.tandai_hadir(poli, loket)
                else:
                    # Pasien hadir — cukup hapus dari cache konfirmasi
                    puskesmas.tandai_hadir(poli, loket)
                    print(f"\n  ✅  {pasien['nama']} tercatat HADIR.")
            jeda()

        elif pilihan == "2":
            # ── MENU [2]: Koreksi tidak hadir untuk pasien yang sudah dipanggil ──
            os.system("cls" if os.name == "nt" else "clear")
            header("⚠️   TANDAI TIDAK HADIR")

            # Cek apakah ada pasien yang masih menunggu konfirmasi di cache
            if not puskesmas.pasien_terakhir_dipanggil:
                print("\n  ℹ  Tidak ada pasien yang perlu dikonfirmasi saat ini.")
                print("     Panggil pasien terlebih dahulu melalui Menu [1].")
                jeda()
                continue

            # Tampilkan daftar pasien yang ada di cache konfirmasi
            print("\n  Pasien yang menunggu konfirmasi kehadiran:\n")
            daftar_kunci = list(puskesmas.pasien_terakhir_dipanggil.keys())
            for idx, (poli, loket) in enumerate(daftar_kunci, start=1):
                pasien = puskesmas.pasien_terakhir_dipanggil[(poli, loket)]
                print(f"  [{idx}] {pasien['nama']:<25} → {poli} Loket {loket}  (No. {pasien['nomor_antrian']})")
            print("  [0] Kembali")

            while True:
                try:
                    pilihan_koreksi = input("\n  Pilih pasien yang tidak hadir: ").strip()
                    if pilihan_koreksi == "0":
                        break
                    nomor_pilihan = int(pilihan_koreksi)
                    if 1 <= nomor_pilihan <= len(daftar_kunci):
                        poli_terpilih, loket_terpilih = daftar_kunci[nomor_pilihan - 1]
                        print()
                        berhasil = puskesmas.tandai_tidak_hadir_dari_cache(poli_terpilih, loket_terpilih)
                        if berhasil:
                            print(f"\n  ✅  Status tidak hadir berhasil dicatat.")
                        break
                    print(f"  ⚠  Pilihan harus antara 1–{len(daftar_kunci)}.")
                except ValueError:
                    print("  ⚠  Masukkan angka yang valid.")
            jeda()

        elif pilihan == "3":
            os.system("cls" if os.name == "nt" else "clear")
            puskesmas.tampilkan_status()
            jeda()

        elif pilihan == "4":
            os.system("cls" if os.name == "nt" else "clear")
            header("📂  DAFTAR ULANG DARI EXCEL")
            if data_antrian_dummy:
                # Daftarkan ulang seluruh data dari Excel ke dalam antrian
                for pasien in data_antrian_dummy:
                    puskesmas.tambah_ke_antrean(pasien)
                print("\n  ✅ Semua pasien dari Excel telah didaftarkan ulang.")
            else:
                print("  ❌ Tidak ada data Excel yang tersedia.")
            jeda()

        elif pilihan == "5":
            os.system("cls" if os.name == "nt" else "clear")
            header("📝  DAFTARKAN PASIEN BARU")
            puskesmas.daftar_pasien_manual()
            jeda()

        elif pilihan == "0":
            print("\n  Sampai jumpa! 👋\n")
            break

        else:
            print("  ⚠  Pilihan tidak valid. Coba lagi.")
            time.sleep(1)
            os.system("cls" if os.name == "nt" else "clear")


# ══════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════
if __name__ == "__main__":
    os.system("cls" if os.name == "nt" else "clear")

    # Inisialisasi sistem dengan 3 loket Poliklinik Umum dan 2 loket Poliklinik Gigi
    puskesmas = SistemAntrianPuskesmas(jumlah_loket_umum=3, jumlah_loket_gigi=2)

    if data_antrian_dummy:
        header("📂  PENDAFTARAN AWAL PASIEN")
        # Daftarkan semua pasien dari Excel ke antrian saat program pertama kali dijalankan
        for pasien in data_antrian_dummy:
            puskesmas.tambah_ke_antrean(pasien)
        jeda()

    # Jalankan loop menu utama
    menu_utama(puskesmas)
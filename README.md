=======================================================
   GENERATOR RISALAH LELANG — PANDUAN PENGGUNAAN
   Versi 1.0
=======================================================

─────────────────────────────────────────────────────
LANGKAH 1 — INSTALL SATU KALI (Hanya pertama kali)
─────────────────────────────────────────────────────

1. Pastikan Python sudah terinstall di komputer Anda.
   Jika belum, download dari: https://www.python.org/downloads/
   ⚠️  Saat install, centang "Add Python to PATH"

2. Buka Command Prompt (tekan tombol Windows, ketik "cmd", Enter)

3. Ketik perintah berikut, tekan Enter:

      pip install openpyxl python-docx

4. Tunggu hingga selesai. Cukup dilakukan sekali.


─────────────────────────────────────────────────────
LANGKAH 2 — CARA MEMBUKA APLIKASI
─────────────────────────────────────────────────────

Klik dua kali file:  app_risalah_lelang.py

  ATAU

Klik kanan → "Open with" → Python


─────────────────────────────────────────────────────
LANGKAH 3 — CARA MENGGUNAKAN APLIKASI
─────────────────────────────────────────────────────

1. Klik [📂 Pilih File] → pilih file Excel laporan lelang
   (biasanya berakhiran .xlsx)

2. Klik [📂 Pilih Folder] → pilih lokasi penyimpanan
   hasil dokumen Word
   (default: folder yang sama dengan file Excel)

3. Isi kolom "Nomor Risalah Lelang" dan "Tanggal Lelang"
   sesuai dokumen lelang saat ini.

4. Pastikan "Nama Pejabat Lelang" sudah benar.

5. Klik [▶ BUAT RISALAH LELANG]

6. Tunggu hingga muncul pesan "Berhasil!"
   Dokumen Word akan tersimpan otomatis.


─────────────────────────────────────────────────────
KETERANGAN PANEL LOG
─────────────────────────────────────────────────────

Panel kanan (Log Proses) menampilkan detail proses:

  ✅  Hijau   = berhasil / informasi penting
  ⚠️  Kuning  = peringatan (data tidak ditemukan, dsb)
  ❌  Merah   = error / gagal

Jika ada error, screenshot panel Log dan kirim ke IT.


─────────────────────────────────────────────────────
PERTANYAAN UMUM
─────────────────────────────────────────────────────

T: "Aplikasi tidak bisa dibuka"
J: Pastikan Python sudah terinstall. Ulangi Langkah 1.

T: "Muncul pesan kolom tidak ditemukan"
J: Pastikan file Excel memiliki sheet "FIRMAN" dengan
   kolom bernama: Lot, Status, Limit, Laku, TAP

T: "Section Laku / TAP kosong"
J: Periksa kolom Status di Excel. Nilai status yang
   dikenali sebagai Laku: Sold, Terjual
   Nilai status yang dikenali sebagai TAP: Not Sold,
   Tak Terjual, Tidak Laku, TAP, Tidak Terjual

T: "Format Rupiah salah"
J: Aplikasi sudah memperbaiki format Rp secara otomatis.
   Jika masih salah, hubungi IT.


─────────────────────────────────────────────────────
KONTAK / BANTUAN
─────────────────────────────────────────────────────

Jika mengalami kendala, hubungi tim IT dengan
menyertakan screenshot panel Log di aplikasi.

=======================================================# lelangApp
generate risalah from sheet to word

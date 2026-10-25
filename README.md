# Aircraft Battery Maintenance Monitoring System (Unit TCC-2)
### PT GMF AeroAsia Tbk — Desktop Dashboard & Data Integrity System

Aplikasi desktop GUI modern, tangguh, dan **100% Berjalan Sepenuhnya Offline** yang dibangun menggunakan **Python**, **CustomTkinter**, **ttk.Treeview**, **pandas**, dan **openpyxl**. Aplikasi ini dirancang untuk memantau pengerjaan dan riwayat perawatan baterai pesawat (Aircraft Battery Shop) pada Unit TCC-2 dengan jaminan integritas data tingkat tinggi.

---

## 📌 1. Prinsip Utama & Arsitektur Sistem

Aplikasi ini beroperasi menggunakan prinsip arsitektur lokal:

$$\textbf{Python Desktop GUI} \longleftrightarrow \textbf{Master Excel Workbook} \longleftrightarrow \textbf{Local Backup Manager}$$

1. **100% Offline**: Tidak memerlukan koneksi internet, LAN bengkel, Wi-Fi, database cloud, atau server eksternal. Semua pemrosesan data dilakukan secara lokal di komputer teknisi.
2. **Satu Sumber Data Resmi (Single Source of Truth)**: File spreadsheet Excel (dinamis sesuai workbook yang di-import atau dibuka pengguna, dengan default `Daily Battery.xlsx`) adalah sumber data resmi yang dapat dibuka dan diedit langsung menggunakan Microsoft Excel.
3. **Penyimpanan Atomik & Aman (Atomic Writing)**: Dashboard tidak pernah menimpa file master secara langsung saat proses penulisan belum terverifikasi. Perubahan ditulis ke file kerja sementara (`.tmp_...`), divalidasi integritasnya, lalu ditukar secara atomik (`os.replace`).
4. **Pencegahan Kehilangan Data (Automated Backups)**: Sebelum modifikasi apa pun dilakukan pada workbook, sistem secara otomatis mengambil snapshot cadangan terverifikasi ke dalam folder lokal `backups/`.
5. **Deteksi Konflik Perubahan Eksternal (Fingerprinting)**: Menggunakan sidik jari berkas kombinasi SHA-256 dan modifikasi timestamp (`mtime`) untuk mendeteksi apakah teknisi telah mengubah file di Microsoft Excel saat dashboard sedang terbuka.
6. **Konsistensi Riwayat (Historical Preservation)**: Sheet `History FSB` dan `History SCRAP` tidak pernah dikorbankan atau dihapus saat pembaruan data harian dilakukan.

---

## 🏗️ 2. Struktur Modul & Pemisahan Tanggung Jawab

Kode aplikasi telah direfaktor secara bertahap ke dalam modul-modul independen untuk kemudahan pemeliharaan dan keandalan tinggi:

```
battery_shop_tcc2/
│
├── app.py                   # Antarmuka Pengguna Desktop (CustomTkinter & ttk.Treeview)
├── excel_repository.py      # Lapisan Data: Pembacaan murni, penulisan atomik, & safe export
├── sync_service.py          # Sinkronisasi: Fingerprinting SHA-256, deteksi konflik, deteksi lock
├── validation.py            # Validasi integritas: Skema workbook, record baterai, & deteksi duplikasi
├── backup_manager.py        # Pengelola backup: Pembuatan snapshot, rotasi retensi, & pemulihan
├── config.py                # Konfigurasi sistem: Lokasi file master, path backup, & retensi
│
├── config.json              # File konfigurasi master path & backup policy (dibuat otomatis)
├── dropdown_config.json     # Opsi dropdown tersimpan & hasil auto-learning
├── audit_log.jsonl          # Catatan audit lokal (waktu, jenis aksi, WO, dan detail)
├── assets/                  # Aset visual & logo resmi PT GMF AeroAsia Tbk (gmf_logo.png)
├── backups/                 # Folder penyimpanan snapshot cadangan berkala
├── Daily Battery.xlsx       # Workbook Excel Master resmi
├── start.bat                # Script peluncur Windows desktop 1-klik (PC/Laptop)
└── test_system_integrity.py # Test suite otomatis integritas data & offline
```

---

## 🔄 3. Alur Kerja Sinkronisasi Data

### A. Sinkronisasi dari Microsoft Excel $\rightarrow$ Dashboard
Ketika teknisi mengubah atau menambahkan data secara langsung menggunakan Microsoft Excel:
1. Teknisi menyimpan perubahan di Microsoft Excel (`Ctrl + S`) dan menutup file (atau membiarkannya tersimpan).
2. Buka aplikasi dashboard Python.
3. Tekan tombol **`🔄 Sync / Refresh`** pada toolbar atas (atau tekan tombol pintasan keyboard **`F5`**).
4. `ExcelRepository` membaca data seluruh sheet secara murni (*read-only*), memvalidasi skema, dan memperbarui memori DataFrame tanpa menulis ulang atau merusak formatting workbook.
5. Tabel, kartu KPI, dropdown, pencarian, dan kalender langsung menampilkan data terkini.
6. Notifikasi sukses menampilkan rincian jumlah unit aktif dan riwayat yang dimuat.

### B. Sinkronisasi dari Dashboard $\rightarrow$ Microsoft Excel
Ketika teknisi mengedit sel langsung (*inline edit*), menambah unit, atau menghapus baris melalui dashboard:
1. **Validasi Input**: Validasi kelengkapan field utama (`WO`, `Serial Number`, `Part Number`, `Customer`, nilai `TAT`).
2. **Pemeriksaan Duplikasi**: Nomor WO diperiksa agar tidak terjadi duplikasi unit aktif pada sheet Daily Battery.
3. **Deteksi Konflik Eksternal**: Jika file Excel master telah dimodifikasi oleh program luar sejak terakhir disinkronkan, dialog peringatan akan muncul menanyakan apakah pengguna ingin membatalkan untuk me-refresh data terlebih dahulu atau melanjutkan penimpaan.
4. **Pemeriksaan File Terkunci**: Jika file sedang dikunci oleh Microsoft Excel, sistem menampilkan instruksi yang ramah agar pengguna menyimpan dan menutup Excel terlebih dahulu, mencegah kegagalan simpan.
5. **Backup Otomatis**: Sistem membuat backup bertanggal di folder `backups/`.
6. **Penulisan Atomik**: Data ditulis ke file sementara, diverifikasi ulang, lalu diganti secara atomik ke file master.
7. **Audit Trail**: Perubahan dicatat secara lokal ke `audit_log.jsonl`.
8. Jika terjadi kegagalan sistem, status in-memory dan tampilan tabel segera di-rollback ke kondisi semula.

---

## 🛡️ 4. Panduan Backup & Pemulihan (Disaster Recovery)

Jika terjadi kesalahan manusia di Microsoft Excel, kerusakan file, atau listrik padam:

### Cara Melakukan Pemulihan (Restore) melalui GUI:
1. Klik status bar bagian bawah: **`● Live Master: Daily Battery.xlsx [Klik untuk Kelola Excel]`**, atau klik tombol **`📑 Kelola Excel`** pada toolbar.
2. Pada dialog yang muncul, pilih tombol **`📦 Pulihkan dari Backup`**.
3. Pilih salah satu titik snapshot yang diinginkan dari daftar cadangan (urut dari yang terbaru).
4. Klik **`Pulihkan Snapshot Terpilih`** dan konfirmasi.
5. Sistem akan mencadangkan kondisi saat ini terlebih dahulu, lalu mengganti master dengan snapshot yang dipilih, dan memuat ulang dashboard secara otomatis.

### Cara Pemulihan Manual dari Folder `backups/`:
1. Buka folder `backups/` di direktori aplikasi.
2. Cari file backup dengan tanggal dan waktu yang diinginkan, misalnya:
   `Daily_Battery_backup_20261015_214544_pre_save_Daily_Battery.xlsx`
3. Salin (*copy*) file tersebut ke folder utama aplikasi dan ubah namanya menjadi `Daily Battery.xlsx`.
4. Buka kembali dashboard dan tekan tombol **`🔄 Sync / Refresh`**.

---

## 📑 5. Membedakan Master Workbook vs Export Workbook

Sistem membedakan secara tegas antara file master dan file ekspor:
- **Master Workbook**: Sumber data resmi aktif yang terikat dengan dashboard (misalnya file Excel yang di-import dari HP atau `Daily Battery.xlsx`). Lokasi dan nama file dapat diganti kapan saja melalui tombol *📁 Excel ▾* pada header atau menu *Pengaturan*.
- **Export Workbook**: Laporan statis yang diekspor untuk distribusi, cetak, atau laporan manajemen. Fitur ekspor di dashboard secara ketat **melarang** pengguna mengekspor ke nama file yang sama dengan file master aktif untuk mencegah penimpaan tidak sengaja.

---

## 🚀 6. Cara Instalasi & Menjalankan (100% Offline)

### Persyaratan Sistem:
- Sistem Operasi: Windows 10 / Windows 11 (64-bit disarankan).
- Python 3.10 atau versi yang lebih baru.
- Pustaka Python:
  ```bash
  pip install customtkinter pandas openpyxl tkcalendar
  ```

### Menjalankan Aplikasi:
Cukup klik dua kali file batch:
```
start.bat
```
Atau jalankan melalui command prompt / PowerShell:
```powershell
python app.py
```

### Menjalankan Pengujian Integritas Mandiri:
Untuk memverifikasi bahwa seluruh mekanisme keamanan data berjalan dengan baik:
```powershell
python test_system_integrity.py
```

---

## 📦 7. Panduan Build Executable Mandiri (PyInstaller)

Jika ingin mendistribusikan aplikasi sebagai file executable tunggal (`.exe`) tanpa perlu menginstal Python di komputer target:

1. Pasang PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Jalankan perintah kompilasi berikut dari folder aplikasi:
   ```powershell
   pyinstaller --noconfirm --onedir --windowed --name "Battery_Shop_TCC2" --add-data "dropdown_config.json;." --add-data "Daily Battery.xlsx;." app.py
   ```
3. Folder hasil build akan tersedia di folder `dist/Battery_Shop_TCC2/`. Seluruh aset, pustaka `customtkinter`, dan template Excel sudah terpaket secara mandiri.

---

## ⚠️ 8. Batasan Penggunaan Satu Komputer (Single-PC Constraints)

1. **Konkurensi**: Karena aplikasi ini menggunakan spreadsheet Excel lokal dan bukan database relasional client-server, aplikasi ditujukan untuk dioperasikan pada satu komputer bengkel utama per unit shop.
2. **Penguncian Berkas Microsoft Excel**: Saat file master sedang diedit di Microsoft Excel, Windows memberlakukan penguncian file eksklusif (*file lock*). Jika teknisi menekan tombol simpan di dashboard saat Excel masih terbuka dan terkunci, dashboard akan memperingatkan teknisi agar menutup file di Excel terlebih dahulu demi menjamin keselamatan data.
3. **Penyimpanan Eksternal**: Jika master workbook disimpan di flashdisk atau harddisk eksternal, pastikan media penyimpanan tidak dicabut saat aplikasi sedang bekerja.

---

## 📋 9. Pintasan Keyboard & Navigasi Penting

- **`F5`**: Sinkronisasi ulang data dari file Excel (*Sync / Refresh Data*).
- **`F1`**: Membuka Glosarium dan Istilah Teknis Battery Shop TCC-2.
- **`Delete`**: Menghapus baris unit baterai yang dipilih (dengan dialog konfirmasi aman).
- **`Double-Click Sel`**: Membuka editor cepat sel langsung pada tabel (dropdown atau input teks).
- **`Shift + Scroll Mouse`**, **Scroll di Status Bar / Slider**, atau **Drag Slider Horizontal**: Menggeser tampilan ke samping (*horizontal slide*) yang mulus dan responsif untuk mengakses seluruh fitur di resolusi layar laptop apa pun.
- **`Slider Vertikal (Scroll Kebawah) pada Dialog`**: Seluruh modal (termasuk fitur *Akses & Manajemen Excel* di pojok kanan atas) kini dilengkapi slider vertikal sehingga seluruh aksi (ekspor, cadangan, restore) selalu dapat diakses penuh tanpa terpotong di layar kecil/High-DPI.

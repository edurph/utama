# 🏫 e-RPH Digital 2026

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)](https://developers.google.com/apps-script)
[![Tailwind CSS](https://img.shields.io/badge/Tailwind_CSS-38B2AC?style=for-the-badge&logo=tailwind-css&logoColor=white)](https://tailwindcss.com/)
[![PWA Ready](https://img.shields.io/badge/PWA-Ready-orange?style=for-the-badge&logo=pwa&logoColor=white)](https://web.dev/progressive-web-apps/)
[![Status](https://img.shields.io/badge/Status-Aktif%20%26%20Stabil-success?style=for-the-badge)]()

Sistem Pengurusan Rekod Rancangan Pengajaran Harian (**e-RPH**) Digital, Pemetaan DSKP Pintar, Penjanaan Draf Pukal Sebulan, dan Kawalan Akses Langganan (**Whitelist**) berasaskan Google Apps Script & Google Sheets.

---

## 🌟 Ciri-Ciri Utama Sistem

### 1. 🛡️ Langkah 1: Profil Guru & Kawalan Akses (Whitelist)
* **Borang Profil Responsif**: Pengisian Nama Penuh Guru, Emel Rasmi/DELIMa, Kod & Nama Sekolah, serta Sesi Persekolahan.
* **Pengesahan Berpusat**: Memadankan emel guru secara langsung dengan tab \Senarai_Whitelist\ di Google Sheets.
* **Kunci Lalai (Default Locked)**: Langkah 2 (Jadual) dan Langkah 3 (Jana RPH) dikunci secara automatik sehingga profil berjaya disahkan.
* **Penyimpanan Sesi Pintar**: Profil yang disahkan disimpan dalam *local storage* pelayar untuk akses segera tanpa pengesahan berulang.

### 2. 📅 Langkah 2: Templat Jadual Waktu & Penjanaan Seminggu
* **Penetapan Jadual Waktu Tetap**: Guru hanya perlu tetapkan jadual sekali, dan sistem menyimpannya ke Google Sheets (\JADUAL_GURU\).
* **Enjin DSKP Terperinci**: Sokongan lengkap bagi Kurikulum Standard Sekolah Rendah (Perdana, DLP, Pemulihan Khas, Prasekolah, dan PPKI).
* **Penjanaan Draf Seminggu**: Bina draf RPH lengkap (SK, SP, Objektif, Kriteria Kejayaan, Aktiviti PAK21, BBM, Nilai, KBAT, Refleksi).
* **Eksport PDF Segera**: Cetak RPH mingguan berformat standard A4 portrait terus ke Google Drive.

### 3. ⚡ Langkah 3: Jana RPH Sebulan (Pukal / Bulk Auto-Generator)
* **Penulisan Berkelompok (Batch Write)**: Menjana rekod 4 hingga 5 minggu dalam satu bulan serentak ke tab \RPH_GURU\ dalam beberapa saat.
* **Paparan Accordion / Kalendar**: Menyemak setiap minggu dan hari dengan kemas dan teratur.
* **Pengisian Refleksi Pukal**: Isi refleksi secara pukal bagi seluruh minggu/bulan dengan satu klik.
* **Kompilasi PDF Sebulan**: Gabungkan seluruh rekod PdP sebulan ke dalam satu fail PDF bersaiz kemas.

---

## 📁 Struktur Fail Projek

\\\	ext
├── index.html                       # Shell PWA Utama untuk GitHub Pages (Root)
├── manifest.json                    # Konfigurasi PWA Web App Manifest
├── sw.js                            # Service Worker untuk pasang di telefon pintar
├── icon-192.png                     # Ikon PWA standard (192x192) dengan Logo Baharu
├── icon-512.png                     # Ikon PWA resolusi tinggi (512x512) dengan Logo Baharu
├── logo_erph_2026.jpg               # Logo rasmi e-RPH baharu resolusi tinggi
├── Code.gs                          # Logik Backend GAS (Whitelist, Profil, Enjin DSKP, PDF)
├── Index_GAS.html                   # Antaramuka Hadapan untuk Google Apps Script
├── Kokurikulum.gs                   # Modul Kehadiran & Aktiviti Kokurikulum
├── PANDUAN_SETUP_GOOGLE_SHEETS.md   # Panduan konfigurasi pangkalan data Sheets
└── apps-script/                     # Salinan khusus untuk Google Apps Script
    ├── Code.gs
    └── Index.html
\\\

---

## 🚀 Panduan Setup Pantas

### Fasa 1: Google Sheets & Apps Script
1. Buka [Google Sheets](https://sheets.new) dan cipta spreadsheet baharu.
2. Buka menu **Extensions (Pelanjutan)** ➔ **Apps Script**.
3. Salin kandungan \Code.gs\ ke dalam Apps Script.
4. Tambah fail HTML bernama \Index.html\ dan salin kandungan \Index_GAS.html\.
5. Klik butang **Deploy** ➔ **New deployment** ➔ Jenis **Web app** (Execute as: \Me\, Who has access: \Anyone\).
6. Salin URL Web App yang berakhir dengan \/exec\.

### Fasa 2: GitHub Pages (PWA App)
1. Buka repositori \https://github.com/edurph/utama\.
2. Muat naik semua fail dalam folder ini ke GitHub.
3. Di tab **Settings** ➔ **Pages**: Tetapkan Branch ke **main** ➔ Klik **Save**.
4. Buka laman web rasmi anda di: **https://edurph.github.io/utama/**.
# ðŸ“Š Panduan Konfigurasi Google Sheets untuk e-RPH Digital 2026 & Mod DFY 40 Minggu

Sistem e-RPH Digital 2026 disambungkan secara langsung dengan Google Sheets sebagai pangkalan data berpusat.

---

## ðŸ› ï¸ Senarai Tab Spreadsheet

### 1. Tab `Senarai_Whitelist` (Kawalan Akses Guru Standard)
Digunakan oleh pentadbir untuk menentukan emel guru yang dibenarkan menggunakan sistem web app mingguan:
| Lajur A | Lajur B | Lajur C | Lajur D | Lajur E |
|---|---|---|---|---|
| **Emel** | **Nama_Penuh** | **Peranan** | **Status** | **Tarikh_Didaftarkan** |
| contoh@moe-dl.edu.my | Cikgu Ali | Guru | AKTIF | 10/09/2026 |

*Nota: Hanya emel berstatus **AKTIF** dibenarkan melepasi Langkah 1.*

---

### 2. Tab `TokenDatabase` (Penebusan DFY 40 Minggu Shopee)
Pangkalan data token keselamatan bagi pembeli Shopee yang membeli pakej DFY (Done-For-You) 40 Minggu:
| Lajur A | Lajur B | Lajur C | Lajur D | Lajur E | Lajur F | Lajur G |
|---|---|---|---|---|---|---|
| **Token** | **Status** | **SubjectAllowed** | **YearAllowed** | **RedeemedByEmail** | **RedeemedAt** | **DriveFileUrl** |
| RPH-2026-A1B2 | ACTIVE | ALL | ALL | | | |
| RPH-2026-MATH | ACTIVE | Matematik | Tahun 1 | | | |
| RPH-2026-JAWI | USED | Pendidikan Islam | Tahun 2 | guru@moe-dl.edu.my | 2026-09-13 14:30 | https://docs.google.com/spreadsheets/d/... |

**Peraturan Status Token:**
- `ACTIVE`: Token sedia untuk ditebus oleh pembeli.
- `USED`: Token telah digunakan. Sistem akan menolak sebarang cubaan tebus semula.
- `PROCESSING`: Token dikunci sementara proses kompilasi 40 minggu sedang berjalan bagi mengelakkan penebusan serentak (race condition).
- `SubjectAllowed` / `YearAllowed`: Boleh diisi `ALL` untuk akses terbuka, atau nama subjek / tahun tertentu mengikut pakej belian di Shopee.

---

### 3. Tab `DSKP` / `PANGKALAN_DSKP` (Kemas Kini Silibus Kurikulum)
Anda boleh menyalin dan menampal (copy-paste) terus data DSKP rasmi KPM ke dalam tab ini. Sistem akan mengutamakan data dari tab ini berbanding templat lalai:
| Lajur A | Lajur B | Lajur C | Lajur D | Lajur E | Lajur F | Lajur G | Lajur H |
|---|---|---|---|---|---|---|---|
| **SUBJEK** | **TAHUN** | **TEMA** | **TAJUK** | **SK** | **SP** | **OBJEKTIF** | **AKTIVITI** |
| Bahasa Melayu | Tahun 1 | Kekeluargaan | Hari Pertama Di Sekolah | 1.1 Mendengar dan memberikan respons | 1.1.1 Mendengar, mengecam, memahami sebutan | Pada akhir PdP, murid dapat... | 1. Penerangan guru<br>2. Latihan kumpulan |

*Ciri-ciri Utama:*
- **Ayat Penuh**: Pastikan lajur SK, SP, Tema dan Tajuk mengandungi ayat penerangan penuh.
- **Pengekalan Jawi**: Untuk subjek Pendidikan Islam dan Bahasa Arab, masukkan teks dalam ejaan Jawi/Arab sebenar. Sistem mengekalkan enkod UTF-8 tanpa herotan.

---

### 4. Tab `Rekod_Profil`
Menyimpan rekod profil guru yang berjaya mengesahkan maklumat mereka di Langkah 1.

### 5. Tab `JADUAL_GURU`
Menyimpan penetapan jadual waktu tetap guru bagi memudahkan penjanaan draf berulang.

### 6. Tab `RPH_GURU`
Menyimpan keseluruhan draf dan rekod PdP harian dan bulanan yang dijana melalui sistem web app mingguan.
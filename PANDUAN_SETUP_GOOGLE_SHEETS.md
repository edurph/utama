# 📊 Panduan Konfigurasi Google Sheets untuk e-RPH Digital 2026

Sistem e-RPH Digital 2026 disambungkan secara langsung dengan Google Sheets sebagai pangkalan data berpusat.

---

## 🛠️ Tab Sheet yang Dibina Secara Automatik
Sistem backend \Code.gs\ telah diprogramkan untuk membina tab berikut secara automatik apabila sistem dijalankan buat kali pertama:

### 1. Tab \Senarai_Whitelist\ (Kawalan Akses Guru)
Digunakan oleh pentadbir untuk menentukan emel guru yang dibenarkan menggunakan sistem:
| Lajur A | Lajur B | Lajur C | Lajur D | Lajur E |
|---|---|---|---|---|
| **Emel** | **Nama_Penuh** | **Peranan** | **Status** | **Tarikh_Didaftarkan** |
| contoh@moe-dl.edu.my | Cikgu Ali | Guru | AKTIF | 10/09/2026 |

*Nota: Hanya emel berstatus **AKTIF** dibenarkan melepasi Langkah 1.*

### 2. Tab \Rekod_Profil\
Menyimpan rekod profil guru yang berjaya mengesahkan maklumat mereka di Langkah 1.

### 3. Tab \JADUAL_GURU\
Menyimpan penetapan jadual waktu tetap guru bagi memudahkan penjanaan draf berulang.

### 4. Tab \RPH_GURU\
Menyimpan keseluruhan draf dan rekod PdP harian dan bulanan yang dijana.
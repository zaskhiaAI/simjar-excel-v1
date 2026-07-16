# 🗓️ SIMJAR: Sistem Informasi Manajemen Jadwal Pembelajaran

[![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://products.office.com/excel)
[![Version](https://img.shields.io/badge/Version-v1.0--Excel-blue?style=for-the-badge)](https://github.com/)

> **"Jadwal Cerdas, Pembelajaran Terstruktur."**
> *Membangun Jadwal yang Efektif, Efisien, dan Terintegrasi.*

---

## 📖 1. Latar Belakang
SMKS PGRI 2 Sidoarjo masih melakukan penyusunan jadwal pelajaran secara manual menggunakan Microsoft Excel. Proses ini membutuhkan waktu yang lama karena harus mempertimbangkan banyak aturan, seperti pembagian beban mengajar guru, guru sertifikasi, wali kelas, pembagian *team teaching*, serta menghindari bentrok jadwal. 

Aplikasi SIMJAR (Versi 1.0 - Microsoft Excel + VBA) dirancang untuk membantu Waka Kurikulum menghasilkan jadwal pelajaran secara otomatis, meminimalkan konflik, serta menghasilkan berbagai laporan administrasi sekolah secara instan.

## 🎯 2. Tujuan Aplikasi
Aplikasi ini dikembangkan sebagai instrumen terpadu untuk:
- Mengelola data guru, mata pelajaran, dan kelas.
- Mengelola distribusi mengajar secara proporsional.
- Menghasilkan jadwal pembelajaran secara otomatis.
- Memvalidasi konflik jadwal secara *real-time*.
- Menghasilkan laporan administrasi kurikulum.

---

## 📂 3. Struktur Workbook
Workbook SIMJAR terdiri dari **15 Sheet** yang terbagi menjadi 4 modul utama:

| Modul | No | Nama Sheet | Keterangan |
| :--- | :---: | :--- | :--- |
| **Menu Utama** | - | `DASHBOARD` | Halaman utama dan navigasi aplikasi |
| **Master Data** | 1 | `Guru` | Data identitas dan status guru |
| | 2 | `Mata Pelajaran` | Data mapel dan kelompoknya |
| | 3 | `Kelas` | Data rombongan belajar |
| | 4 | `Distribusi Mengajar` | Inti pemetaan beban mengajar |
| | 5 | `Preferensi Guru` | Aturan hari libur/off guru |
| | 6 | `Konfigurasi` | Parameter dasar penjadwalan |
| **Proses (Engine)** | 7 | `Jadwal Utama` | Area *generate* jadwal otomatis |
| | 8 | `Validasi` | Pengecekan bentrok jadwal |
| | 9 | `Log Generate` | Rekaman proses untuk *debugging* |
| **Output (Laporan)**| 10 | `Jadwal Guru (List)` | Jadwal harian bentuk daftar |
| | 11 | `Jadwal Guru (Matriks)`| Jadwal mingguan bentuk tabel |
| | 12 | `Rekap Beban Mengajar` | Rincian total JP per guru |
| | 13 | `Rekap Tugas Mengajar` | Format SK Pembagian Tugas |
| | 14 | `Jadwal Kelas` | Jadwal yang ditempel di kelas (Nama Mapel & Guru) |
| | 15 | `Jadwal Kelas (Kode)` | Jadwal dengan kode guru untuk analisis Waka Kurikulum |

---

## 🖥️ 4. Dashboard (Antarmuka Utama)
Merupakan pusat kontrol (Control Panel) aplikasi yang memuat informasi global dan tombol eksekusi:

**Panel Informasi:**
- Tahun Pelajaran & Semester
- Jumlah Guru, Kelas, dan Mata Pelajaran
- Status Generate (Sudah/Belum)

**Panel Tombol (Aksi):**
- Akses ke seluruh sheet Master Data
- Eksekusi Proses: *Validasi Data, Generate Jadwal, Optimasi Jadwal, Cek Konflik, Reset*
- Eksekusi Cetak: *Cetak per jenis laporan atau Cetak Semua Laporan*

---

## 💾 5. Struktur Master Data

### 5.1 Sheet Guru
| Kode | Nama Guru | Status | Wali | Kelas Wali | Max JP/Hari |
| :---: | :--- | :--- | :---: | :---: | :---: |
| 23 | Agus | Serdik | Ya | XI RPL 1 | 8 |
| 24 | Rina | Non Serdik| Tidak | - | 10 |

### 5.2 Sheet Mata Pelajaran
| Kode | Nama Mata Pelajaran | Kelompok |
| :---: | :--- | :--- |
| MP01 | Basis Data | Produktif |
| MP02 | Pemrograman Web | Produktif |
| MP03 | IPAS | Umum |

### 5.3 Sheet Kelas
| Kode | Tingkat | Jurusan | Rombel |
| :---: | :---: | :--- | :--- |
| XRPL1 | X | RPL | 1 |
| XAKL1 | X | AKL | 1 |

### 5.4 Sheet Distribusi Mengajar (Core Data)
Satu guru dapat mengajar >1 mapel; Satu mapel dapat diajar >1 guru.
| Kode Guru | Mapel | Tingkat | Jurusan | Rombel | JP |
| :---: | :--- | :---: | :--- | :--- | :---: |
| 23 | Basis Data | XI | RPL | 1 | 6 |
| 23 | Pemrograman Web| XI | RPL | 1 | 4 |
| 18 | IPAS | X | AKL | 1 | 4 |

### 5.5 Sheet Preferensi Guru (Soft Constraints)
Guru Serdik dikosongkan.
| Kode Guru | Libur 1 | Libur 2 | Prioritas |
| :---: | :--- | :--- | :--- |
| 24 | Senin | Selasa | Tinggi |
| 25 | Rabu | - | Sedang |

### 5.6 Sheet Konfigurasi
| Parameter | Nilai |
| :--- | :---: |
| Hari Aktif | 5 |
| JP per Hari | 10 |
| Max Generate | 500 |
| Max JP Serdik | 8 |

---

## ⚙️ 6. Proses Penjadwalan (Scheduling Engine)

- **Jadwal Utama:** Menyajikan hasil *generate* di mana sel berisi **Kode Guru**.
- **Validasi:** Menampilkan daftar *conflict* (Contoh: Guru mengajar dua kelas bersamaan, Team Teaching di hari yang sama, dll). Jika kosong, jadwal dinyatakan 100% Valid.
- **Log Generate:** Merekam langkah algoritma (mulai, penempatan guru, perpindahan sel, hingga selesai) untuk keperluan evaluasi/debugging.

---

## 📊 7. Output & Pelaporan

SIMJAR menghasilkan 6 jenis laporan otomatis:
1. **Jadwal Guru (List):** Menampilkan rincian hari, jam, mapel, dan kelas ke bawah. Sangat mudah dibaca guru.
2. **Jadwal Guru (Matriks):** Berbentuk tabel matriks (Hari vs Jam) agar guru mudah melihat jadwal kosong.
3. **Rekap Beban Mengajar Guru:** Dicetak per guru (1 halaman) berisi rincian mapel yang diajar dan total JP.
4. **Rekap Pembagian Tugas Mengajar:** Tabel komprehensif seluruh guru beserta rincian mengajar di setiap kelas. *Sangat ideal untuk lampiran SK Mengajar.*
5. **Jadwal Kelas (Siswa):** Menampilkan nama mapel dan nama/kode guru untuk dipasang di dalam ruang kelas.
6. **Jadwal Kelas (Kode Guru):** Tampilan khusus Waka Kurikulum untuk analisis distribusi kode guru.

---

## ⚖️ 8. Rules & Constraint (Aturan Logika Penjadwalan)

### 🔴 Hard Constraint (Wajib Terpenuhi)
1. Guru tidak boleh mengajar di dua kelas pada waktu (*jam*) yang sama.
2. Satu kelas hanya boleh memiliki satu mata pelajaran pada satu waktu.
3. Guru hanya mengajar sesuai dengan data **Distribusi Mengajar**.
4. Total JP setiap guru harus terpenuhi dengan presisi.
5. Wali kelas **wajib** mengajar di kelas perwaliannya (jika memiliki alokasi jam).
6. *Team Teaching* pada kelas yang sama harus dijadwalkan pada hari yang berbeda.
7. *Team Teaching* tidak boleh berada pada jam yang sama dalam satu hari.
8. Guru **Serdik** wajib memiliki minimal 1 (satu) jam mengajar di setiap hari kerja aktif.

### 🟡 Soft Constraint (Diusahakan Terpenuhi)
1. Guru **Non Serdik** dapat memilih maksimal 2 preferensi hari libur.
2. Jika memungkinkan, jadwal guru Non Serdik dipadatkan agar memperoleh hari kosong (libur).
3. Mengurangi "jam kosong" di tengah-tengah rentang jadwal harian guru.
4. Menyebarkan beban mengajar harian secara proporsional agar guru tidak *overload* di satu hari.
5. *Catatan:* Preferensi ini dapat diabaikan oleh sistem apabila pemenuhan *Hard Constraint* terancam gagal.

---

## 🔄 9. Alur Penggunaan Aplikasi (Workflow)

```mermaid
graph TD
    A[1. Input Master Guru] --> B[2. Input Master Mapel]
    B --> C[3. Input Master Kelas]
    C --> D[4. Input Distribusi Mengajar]
    D --> E[5. Input Preferensi Guru]
    E --> F[6. Validasi Data Awal]
    F --> G{7. Generate Jadwal}
    G --> H[8. Optimasi Jadwal]
    H --> I[9. Cek Konflik]
    I -->|Ada Konflik| G
    I -->|Valid 100%| J[10. Cetak Seluruh Laporan]

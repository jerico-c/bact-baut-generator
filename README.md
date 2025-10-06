# BACT & BAUT Document Generator

Sebuah aplikasi web sederhana untuk mengotomatiskan pembuatan dokumen **Berita Acara Commissioning Test (BACT)** dan **Berita Acara Uji Terima (BAUT)**. Aplikasi ini dibuat untuk mempercepat alur kerja dengan input data dan gambar yang fleksibel.

## ✨ Fitur Utama

  - **Generate Dokumen `.docx`**: Menghasilkan file BACT dan BAUT secara terpisah berdasarkan template.
  - **Input Gambar Fleksibel**: Mendukung *copy-paste* langsung dari clipboard (misalnya dari aplikasi screenshot) dan *drag-and-drop* file gambar.
  - **Input BoQ Cepat**: Mendukung pengisian tabel *Bill of Quantity* dengan *copy-paste* data dari Excel dan navigasi keyboard antar kolom.
  - **Kalkulasi Otomatis**: Menghitung nilai "Tambah" dan "Kurang" pada tabel BoQ secara otomatis di sisi server.
  - **Laporan Gambar Kurang**: Menampilkan daftar gambar yang belum diisi untuk verifikasi cepat sebelum men-download dokumen.

## 💻 Teknologi yang Digunakan

  * **Backend**: Node.js, Express.js, Multer, Docxtemplater
  * **Frontend**: HTML, CSS, Vanilla JavaScript (tanpa framework)

## 📋 Prasyarat

Sebelum memulai, pastikan di komputer Anda sudah terinstal:

  - **Node.js**: Versi 16 atau yang lebih baru direkomendasikan. Anda bisa mengunduhnya di [nodejs.org](https://nodejs.org/).

-----

## 🚀 Instalasi & Penyiapan

Ikuti langkah-langkah berikut untuk menjalankan aplikasi ini di komputer Anda.

### Langkah 1: Unduh Kode Proyek

1.  Buka browser dan kunjungi alamat [https://github.com/jerico-c/bact-baut-generator](https://github.com/jerico-c/bact-baut-generator).

2.  Klik tombol hijau **`< > Code`**, lalu pilih **`Download ZIP`**.

3.  Setelah file ZIP terunduh, **ekstrak** isinya ke lokasi yang Anda inginkan di komputer Anda.

### Langkah 2: Siapkan File Template

  - Pastikan Anda memiliki folder `templates` di dalam direktori proyek.
  - Letakkan file template `BACT_Template.docx` dan `BAUT_Template.docx` yang sudah benar di dalam folder `templates` tersebut.

### Langkah 3: Install Dependensi

1.  Buka **terminal** atau **Command Prompt** di dalam folder proyek Anda yang sudah diekstrak.
2.  Jalankan perintah berikut untuk menginstal semua *library* yang dibutuhkan. Proses ini mungkin memakan waktu beberapa saat.
    ```bash
    npm install
    ```

### Langkah 4: Jalankan Server

Setelah instalasi selesai, jalankan server dengan perintah:

```bash
node server.js
```

Jika berhasil, Anda akan melihat pesan `🚀 Server berjalan di http://localhost:3000` di terminal Anda. Biarkan terminal ini tetap terbuka selama Anda menggunakan aplikasi.

-----

## 📝 Cara Penggunaan Aplikasi

1.  Buka browser (Chrome, Firefox, dll.) dan kunjungi alamat `http://localhost:3000`.
2.  **Isi Nama LOP**: Ketik atau *paste* nama LOKASI proyek pada kolom yang tersedia.
3.  **Isi Tabel BoQ (untuk BAUT)**:
      * Ketik manual nilai "Kontrak" dan "Aktual". Gunakan tombol panah untuk berpindah.
      * Atau, salin (copy) data dari Excel lalu *paste* (Ctrl+V) di salah satu sel untuk mengisinya secara otomatis.
4.  **Masukkan Gambar Bukti**:
      * Klik salah satu kotak gambar untuk mengaktifkannya (akan muncul sorotan biru).
      * **Cara 1 (Paste):** Langsung tekan `Ctrl+V` atau `Cmd+V` untuk menempelkan gambar dari clipboard Anda.
      * **Cara 2 (Drag & Drop):** Seret file gambar dari komputer Anda dan lepaskan di dalam kotak yang aktif.
5.  **Generate Dokumen**:
      * Klik tombol **Generate BACT** atau **Generate BAUT**.
      * Daftar gambar yang kurang (jika ada) akan muncul di bagian bawah. Anda bisa menyalin daftar ini dengan tombol "Copy".
      * File `.docx` yang sudah jadi akan otomatis terunduh ke komputer Anda.

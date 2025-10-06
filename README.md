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

  - **Node.js**: Versi 16 atau yang lebih baru direkomendasikan.

### Panduan Instalasi Node.js untuk Pemula

Jika Anda belum pernah menginstal Node.js, ikuti langkah-langkah mudah berikut:

1.  **Kunjungi Situs Resmi Node.js**

      * Buka browser Anda dan pergi ke [https://nodejs.org/](https://nodejs.org/).

2.  **Pilih Versi yang Tepat**

      * Di halaman utama, Anda akan melihat dua pilihan unduhan. Pilih versi **LTS** (Long Term Support). Ini adalah versi yang paling stabil dan direkomendasikan untuk sebagian besar pengguna.

3.  **Unduh dan Jalankan Installer**

      * Situs web akan secara otomatis mendeteksi sistem operasi Anda (Windows atau Mac). Klik tombol LTS untuk mengunduh file installer (`.msi` untuk Windows, `.pkg` untuk Mac).
      * Setelah unduhan selesai, buka file tersebut.
      * Ikuti petunjuk instalasi yang muncul di layar. Anda tidak perlu mengubah pengaturan apa pun, cukup klik **"Next"** atau **"Continue"** sampai proses instalasi selesai.

4.  **Verifikasi Instalasi**

      * Setelah instalasi selesai, buka **terminal** baru (di Mac) atau **Command Prompt** baru (di Windows).
      * Ketik perintah berikut dan tekan Enter untuk memeriksa versi Node.js:
        ```bash
        node -v
        ```
      * Jika instalasi berhasil, akan muncul nomor versi seperti `v18.17.1` atau yang lebih baru.
      * Ketik juga perintah berikut untuk memeriksa versi `npm` (yang terinstal bersama Node.js):
        ```bash
        npm -v
        ```
      * Jika ini juga menampilkan nomor versi, berarti Node.js sudah siap digunakan\!

-----

## 🚀 Instalasi & Penyiapan Aplikasi

Setelah Node.js terinstal, ikuti langkah-langkah berikut untuk menjalankan aplikasi ini.

### Langkah 1: Unduh Kode Proyek

1.  Buka browser dan kunjungi alamat [https://github.com/jerico-c/bact-baut-generator](https://github.com/jerico-c/bact-baut-generator).
2.  Klik tombol hijau **`< > Code`**, lalu pilih **`Download ZIP`**.
3.  Setelah file ZIP terunduh, **ekstrak** isinya ke lokasi yang Anda inginkan di komputer Anda.

### Langkah 2: Install Dependensi

1.  Buka **terminal** atau **Command Prompt** di dalam folder proyek Anda yang sudah diekstrak.
2.  Jalankan perintah berikut untuk menginstal semua *library* yang dibutuhkan. Proses ini mungkin memakan waktu beberapa saat.
    ```bash
    npm install
    ```

### Langkah 3: Jalankan Server

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

## Convert PDF DPT ke format Excel
- untuk format pemilu/pilkada 2024 dengan model A-KabKo Daftar Pemilih

## Cara Penggunaan
- buat folder `pdf-resources` dan masukkan file pdf yang akan di convert dengan struktur folder
 ```bash
    pdf-resources
    ├───FOLDER NAMA PROVINSI
    │   ├───FOLDER NAMA KABUPATEN
    │   │   ├───LIST FILE PDF (Tidak masalah jika ada sub folder lain semacam kecamatan,kelurahan,tps)
```
- jalankan `python main.py` (proses offline) atau `python main_ilovepdf.py` jika ingin menggunakan layanan ilovepdf (secara online), dan tunggu hingga selesai
- hasil convert berupa csv dan excel akan tersimpan di folder `results/`
- bisa juga gunakan arguments
  - `--source` path_to_pdf_sources
  - `--results` path_to_results
  - `--deleteOriginal` false/true (jika ingin menghapus file pdf dari source jika berhasil convert)

## Gagal Convert
- jika ada file yang gagal convert, maka file yang gagal akan di copy ke folder `results/error-NAMA KABUPATEN/`
- versi `main.py` kadang tidak bisa mendapatkan kolom RT/RW, sedangkan versi `main_ilovepdf.py` akurat tapi bisa kena timeout dari ilovepdf
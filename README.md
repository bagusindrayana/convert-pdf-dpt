## Convert PDF DPT ke format Excel
- untuk format pemilu 2024 dengan model A-KabKo Daftar Pemilih

## Cara Penggunaan
- buat folder `pdf-resources` dan masukkan file pdf yang akan di convert dengan struktur folder
 ```bash
    pdf-resources
    ├───FOLDER NAMA PROVINSI
    │   ├───FOLDER NAMA KABUPATEN
    │   │   ├───LIST FILE PDF (Tidak masalah jika ada sub folder lain semacam kecamatan,kelurahan,tps)
```
- jalankan `python main.py` dan tunggu hingga selesai
- hasil convert berupa csv dan excel akan tersimpan di folder `results/`

## Gagal Convert
- jika ada file yang gagal convert, maka file yang gagal akan di copy ke folder `results/error-NAMA KABUPATEN/`
import csv
from collections import defaultdict
import glob
import pandas as pd

# Fungsi untuk membaca setiap CSV dan menggabungkan datanya
def buat_report(direktori_csv):
    # Menggunakan dictionary yang menampung set tuple (RT, RW)
    data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(set)))))

    # Mencari semua file CSV di dalam direktori
    csv_files = glob.glob(f"{direktori_csv}/*.csv")
    
    for csv_file in csv_files:
        with open(csv_file, mode='r', newline='', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['kecamatan'] == "KECAMATAN" and row['kelurahan_desa'] == "KELURAHAN":
                    continue
                provinsi = row['provinsi']
                kabupaten = row['kabupaten_kota']
                kecamatan = row['kecamatan']
                kelurahan = row['kelurahan_desa']
                tps = row['nomor_tps']
                rt = row['rt']
                rw = row['rw']

                rt = rt if rt != None and rt != '' and rt != '0' else "000"
                rw = rw if rw != None and rw != '' and rw != '0' else "000"

                # Menyimpan kombinasi RT dan RW di dalam set berdasarkan provinsi, kabupaten, kecamatan, kelurahan, dan TPS
                data[provinsi][kabupaten][kecamatan][kelurahan][tps].add((rt, rw))

    return data

# Fungsi untuk menyimpan report utama ke dalam worksheet
def simpan_report_ke_excel(data, output_excel):
    # List untuk menampung data report yang akan disimpan ke Excel
    report_data = []
    tps_count_data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(int))))

    # Iterasi data yang sudah dikumpulkan
    for provinsi, kabupaten_data in data.items():
        for kabupaten, kecamatan_data in kabupaten_data.items():
            for kecamatan, kelurahan_data in kecamatan_data.items():
                for kelurahan, tps_data in kelurahan_data.items():
                    for tps, rt_rw_set in tps_data.items():
                        # Buat list RT/RW dalam format yang rapi
                        rt_rw_list = ', '.join([f"RT {rt}/RW {rw}" for rt, rw in sorted(rt_rw_set)])

                        # Menyimpan data ke dalam list untuk nantinya di-export ke Excel
                        report_data.append({
                            "NAMA PROVINSI": provinsi,
                            "NAMA KOTA": kabupaten,
                            "NAMA KECAMATAN": kecamatan,
                            "NAMA KELURAHAN": kelurahan,
                            "TPS": tps,
                            "LIST RT/RW": rt_rw_list
                        })

                        # Menghitung jumlah TPS per kelurahan
                        tps_count_data[provinsi][kabupaten][kecamatan][kelurahan] += 1

    # Membuat DataFrame dari report_data
    df_report = pd.DataFrame(report_data)
    
    # List untuk menampung jumlah TPS per kelurahan
    tps_count_report = []
    for provinsi, kabupaten_data in tps_count_data.items():
        for kabupaten, kecamatan_data in kabupaten_data.items():
            for kecamatan, kelurahan_data in kecamatan_data.items():
                for kelurahan, tps_count in kelurahan_data.items():
                    tps_count_report.append({
                        "NAMA PROVINSI": provinsi,
                        "NAMA KOTA": kabupaten,
                        "NAMA KECAMATAN": kecamatan,
                        "NAMA KELURAHAN": kelurahan,
                        "JUMLAH TPS": tps_count
                    })

    # Membuat DataFrame dari tps_count_report
    df_tps_count = pd.DataFrame(tps_count_report)

    # Menulis kedua DataFrame ke dalam file Excel pada worksheet yang berbeda
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_report.to_excel(writer, sheet_name='Laporan TPS', index=False)
        df_tps_count.to_excel(writer, sheet_name='Jumlah TPS per Kelurahan', index=False)

    print(f"Laporan berhasil disimpan ke {output_excel} dengan dua worksheet.")

# Path ke direktori tempat file CSV berada
direktori_csv = './results/all'

# Path output untuk file Excel
output_excel = './results/list_tps_per_rt_rw.xlsx'

# Membuat report dari semua CSV di dalam direktori
data = buat_report(direktori_csv)

# Menyimpan report ke dalam file Excel dengan worksheet tambahan
simpan_report_ke_excel(data, output_excel)

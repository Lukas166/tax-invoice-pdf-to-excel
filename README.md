# Faktur Pajak Streamlit

Aplikasi berbasis Streamlit untuk mengekstrak data dari banyak file PDF Faktur Pajak sekaligus dan mengubahnya menjadi file Excel atau CSV yang siap disalin ke Rekapan Faktur Pajak Keluaran.

Program dibuat untuk Faktur Pajak dengan struktur PDF yang sama seperti contoh yang digunakan. Isi faktur seperti nama customer, jumlah barang, jenis barang, harga, satuan, DPP, PPN, dan PPnBM dapat berbeda-beda karena proses ekstraksi tidak bergantung pada data customer tertentu.

## Fitur Utama

- Bulk upload banyak PDF Faktur Pajak sekaligus.
- Setiap barang dalam faktur diubah menjadi satu baris.
- Mendukung faktur dengan satu barang maupun banyak barang.
- Membuat baris subtotal untuk faktur dengan banyak barang.
- Tidak menggunakan alias atau mapping nama customer.
- Nama customer diambil langsung dari PDF.
- Jenis barang diambil secara lengkap sampai sebelum baris harga, quantity, dan satuan.
- Satuan barang diparsing secara dinamis.
- Mendukung berbagai unit seperti:
  - PCS
  - UNIT
  - SET
  - BOX
  - PACK
  - KG
  - GRAM
  - TON
  - LITER
  - ML
  - METER
  - M2
  - M3
  - ROLL
  - LEMBAR
  - BATANG
  - PASANG
  - LUSIN
  - PALLET
  - DRUM
  - KARUNG
  - LOT
- Unit yang belum terdaftar tetap digunakan dalam bentuk huruf kapital.
- Membaca Potongan Harga per item dan total faktur.
- Membaca PPnBM per item dan total faktur.
- Membaca status PPN tidak dipungut.
- Mengalokasikan DPP, PPN, dan PPnBM ke setiap item secara proporsional.
- Menyesuaikan selisih pembulatan agar subtotal sama dengan nilai pada PDF.
- Mendeteksi nomor faktur duplikat.
- Menampilkan warning jika nilai item berbeda dengan ringkasan faktur.
- Pemilihan kolom output menggunakan checkbox.
- Seluruh kolom aktif secara default.
- Mendukung download Excel dan CSV.
- Hasil Excel dimulai dari sel `B1`.

## Data yang Diekstrak

Program menghasilkan kolom berikut:

| Posisi Excel | Nama Kolom | Sumber Data |
|---|---|---|
| B | TGL | Tanggal tanda tangan pada bagian bawah Faktur Pajak |
| C | NO. FAKTUR PAJAK | Kode dan Nomor Seri Faktur Pajak |
| D | NAMA CUSTOMER | Nama pada bagian Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak |
| E | JENIS BARANG | Seluruh deskripsi barang sebelum baris harga, quantity, dan satuan |
| F | QTY | Jumlah barang dari pola `Rp harga x quantity satuan` |
| G | SATUAN | Satuan barang dari pola `Rp harga x quantity satuan` |
| H | @ RP | Harga satuan barang |
| I | DPP | Dasar Pengenaan Pajak |
| J | PPN | Pajak Pertambahan Nilai |
| K | JUMLAH | DPP + PPN + PPnBM jika opsi PPnBM diaktifkan |

Kolom A sengaja dikosongkan agar hasil dimulai dari kolom B seperti format Rekapan Faktur Pajak.

## Contoh Jenis Barang

Deskripsi barang tidak dipotong hanya pada nama awal.

Contoh data pada PDF:

```text
POT 40CC LENGKAP TUTUP SEBANYAK 21.000
(DUA PULUH SATU RIBU) PIECE DENGAN HARGA JUAL
SEBESAR RP 621,62 PER PIECE
```

Hasil pada kolom `JENIS BARANG`:

```text
POT 40CC LENGKAP TUTUP SEBANYAK 21.000 (DUA PULUH SATU RIBU) PIECE DENGAN HARGA JUAL SEBESAR RP 621,62 PER PIECE
```

## Struktur Project

```text
FAKTUR_PAJAK_STREAMLIT/
├── .streamlit/
├── .gitignore
├── app.py
├── exporter.py
├── parser.py
├── README.md
└── requirements.txt
```

Keterangan file:

### `app.py`

Berisi tampilan dan alur utama aplikasi Streamlit:

- Bulk upload PDF.
- Pengaturan proses ekstraksi.
- Pemilihan kolom.
- Preview hasil.
- Informasi file yang gagal.
- Peringatan validasi.
- Download Excel.
- Download CSV.

### `parser.py`

Berisi seluruh proses ekstraksi dan perhitungan:

- Membaca text layer PDF.
- Mengambil nomor Faktur Pajak.
- Mengambil tanggal.
- Mengambil nama customer.
- Memisahkan setiap item barang.
- Mengambil deskripsi barang secara lengkap.
- Mengambil quantity, satuan, dan harga satuan.
- Membaca nilai bruto.
- Membaca Potongan Harga.
- Membaca DPP.
- Membaca PPN.
- Membaca PPnBM.
- Membaca status PPN tidak dipungut.
- Mengalokasikan nilai pajak per item.
- Membuat baris item, subtotal, dan pemisah.

### `exporter.py`

Berisi proses pembuatan file Excel dan CSV:

- Menulis header mulai dari `B1`.
- Mengatur urutan kolom.
- Mengatur lebar kolom.
- Mengatur tinggi baris berdasarkan panjang deskripsi.
- Mengatur font dan border.
- Mengatur format tanggal.
- Mengatur format angka dua digit desimal.
- Membuat header dengan filter.
- Membuat subtotal dalam format bold.
- Menjaga nomor faktur sebagai teks agar angka nol di depan tidak hilang.

### `requirements.txt`

Berisi dependency Python yang diperlukan untuk menjalankan aplikasi.

## Persyaratan

Gunakan Python 3.10 atau versi yang lebih baru.

Periksa versi Python:

```bash
python --version
```

Pada beberapa komputer Windows, perintah Python dapat menggunakan:

```bash
py --version
```

## Instalasi

Buka terminal pada folder project.

### 1. Membuat virtual environment

Windows PowerShell:

```powershell
python -m venv .venv
```

Atau:

```powershell
py -m venv .venv
```

### 2. Mengaktifkan virtual environment

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Windows Command Prompt:

```cmd
.venv\Scripts\activate
```

macOS atau Linux:

```bash
source .venv/bin/activate
```

### 3. Menginstal dependency

```bash
pip install -r requirements.txt
```

## Menjalankan Aplikasi

Jalankan perintah berikut dari folder project:

```bash
streamlit run app.py
```

Jika perintah `streamlit` tidak dikenali, gunakan:

```bash
python -m streamlit run app.py
```

Aplikasi akan terbuka melalui browser, biasanya pada:

```text
http://localhost:8501
```

## Cara Penggunaan

1. Jalankan aplikasi Streamlit.
2. Upload satu atau banyak PDF Faktur Pajak.
3. Tentukan pengaturan ekstraksi melalui sidebar.
4. Pilih kolom yang ingin dimasukkan ke hasil.
5. Periksa preview hasil.
6. Periksa bagian file gagal atau peringatan validasi.
7. Download hasil sebagai Excel atau CSV.
8. Salin hasil tersebut ke Rekapan Faktur Pajak utama.

## Pengaturan Aplikasi

### Baris subtotal untuk faktur multi-barang

Jika diaktifkan, faktur dengan lebih dari satu barang akan memiliki satu baris tambahan yang berisi:

- Total DPP.
- Total PPN.
- Total JUMLAH.

Pada baris item, kolom `JUMLAH` dikosongkan agar mengikuti format Rekapan Faktur Pajak.

### Baris kosong antar faktur

Menambahkan satu baris kosong setelah satu faktur selesai agar hasil lebih mudah dibaca dan disalin.

### PPN tidak dipungut menjadi 0.00

Jika pada PDF terdapat keterangan:

```text
Pajak Pertambahan Nilai tidak dipungut
```

maka nilai output PPN dibuat menjadi `0.00`.

Nilai PPN asli dari PDF tetap dibaca dan dapat ditampilkan sebagai warning validasi.

### Sertakan PPnBM dalam JUMLAH

Jika diaktifkan:

```text
JUMLAH = DPP + PPN + PPnBM
```

Jika dinonaktifkan:

```text
JUMLAH = DPP + PPN
```

### Format nomor faktur lama 16 digit

Nomor faktur lama seperti:

```text
0100012467950711
```

akan diubah menjadi:

```text
010.001-24.67950711
```

Nomor Faktur Pajak Coretax dengan panjang berbeda tidak akan diubah.

### Urutkan berdasarkan tanggal dan nomor

Mengurutkan hasil berdasarkan:

1. Tanggal Faktur Pajak.
2. Nomor Faktur Pajak.
3. Nama file PDF.

### Lewati nomor faktur duplikat

Jika terdapat dua PDF dengan nomor faktur yang sama, hanya faktur pertama yang diproses.

File berikutnya akan ditampilkan pada bagian file gagal atau dilewati.

### Pilihan kolom

Seluruh kolom aktif secara default.

Checkbox `Aktifkan semua kolom` dapat digunakan untuk:

- Mengaktifkan seluruh kolom.
- Menonaktifkan seluruh kolom.
- Memilih hanya kolom tertentu.

Urutan output tetap mengikuti urutan standar Rekapan Faktur Pajak.

## Aturan Perhitungan

### Nilai item bruto

Program membaca nilai item yang tersedia pada PDF.

Jika nilai item tidak ditemukan, program menghitung:

```text
Nilai Item = QTY × Harga Satuan
```

### Nilai netto item

```text
Nilai Netto Item = Nilai Item − Potongan Harga Item
```

### Alokasi DPP

Total DPP dari PDF dibagi ke setiap item berdasarkan proporsi nilai netto item.

```text
Proporsi Item = Nilai Netto Item ÷ Total Nilai Netto
```

```text
DPP Item = Total DPP × Proporsi Item
```

### Alokasi PPN

Total PPN dibagi berdasarkan proporsi DPP setiap item.

```text
PPN Item = Total PPN × Proporsi DPP Item
```

### Alokasi PPnBM

Jika jumlah PPnBM per item sama dengan total PPnBM pada ringkasan PDF, program menggunakan nilai PPnBM masing-masing item.

Jika berbeda, total PPnBM dialokasikan berdasarkan proporsi nilai netto item.

### Penyesuaian pembulatan

Nilai DPP, PPN, dan PPnBM dibulatkan ke Rupiah penuh.

Selisih pembulatan ditempatkan pada item terakhir sehingga:

```text
Total DPP Item = DPP PDF
```

```text
Total PPN Item = PPN Output
```

```text
Total PPnBM Item = PPnBM PDF
```

## Format Excel

Hasil Excel memiliki format berikut:

- Nama sheet: `REKAP FAKTUR`.
- Header dimulai dari `B1`.
- Kolom A dikosongkan.
- Font menggunakan `Franklin Gothic Book`.
- Header menggunakan latar abu-abu.
- Header memiliki filter.
- Nomor faktur disimpan sebagai teks.
- Deskripsi barang menggunakan wrap text.
- Subtotal menggunakan tulisan bold.
- Tanggal menggunakan format:

```text
dd-mmm-yy
```

Contoh:

```text
02-Jan-24
```

Angka disimpan sebagai angka Excel dengan dua digit desimal.

Contoh pada regional English:

```text
1,020.00
675.68
689,194.00
75,811.00
4,071,504.00
```

Contoh pada regional Indonesia:

```text
1.020,00
675,68
689.194,00
75.811,00
4.071.504,00
```

Perbedaan titik dan koma mengikuti regional setting Microsoft Excel. Nilai yang disimpan tetap merupakan angka dan dapat digunakan untuk formula Excel.

## Format PDF yang Didukung

Program ditujukan untuk PDF Faktur Pajak yang memiliki struktur tetap dengan keyword seperti:

```text
Kode dan Nomor Seri Faktur Pajak
```

```text
Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak
```

```text
Nama Barang Kena Pajak / Jasa Kena Pajak
```

```text
Rp harga x quantity satuan
```

```text
Potongan Harga = Rp
```

```text
PPnBM (...) = Rp
```

```text
Harga Jual / Penggantian / Uang Muka / Termin
```

```text
Dasar Pengenaan Pajak
```

```text
Jumlah PPN (Pajak Pertambahan Nilai)
```

```text
Jumlah PPnBM (Pajak Penjualan atas Barang Mewah)
```

Data di dalam PDF boleh berbeda-beda selama struktur dan keyword tersebut tetap sama.

## Batasan

- PDF harus memiliki text layer.
- PDF hasil scan berupa gambar belum dapat diproses tanpa OCR.
- Struktur tabel dan keyword PDF harus sama dengan format Faktur Pajak yang didukung.
- PDF dengan layout berbeda mungkin memerlukan penyesuaian parser.
- Program tidak mengubah file Rekapan Faktur Pajak utama secara langsung.
- Program menghasilkan file perantara yang dapat diperiksa dan disalin ke rekapan utama.

## Dependency

Isi minimal `requirements.txt`:

```text
streamlit>=1.40,<2.0
PyMuPDF>=1.24,<2.0
XlsxWriter>=3.2,<4.0
```

Instal seluruh dependency dengan:

```bash
pip install -r requirements.txt
```

## Menghentikan Aplikasi

Pada terminal tempat Streamlit berjalan, tekan:

```text
Ctrl + C
```
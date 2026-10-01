# Cari Barang SSH, SBU, dan ASB Kota Malang

Website statis untuk mencari nama barang, kode rekening, dan harga satuan dari
empat sumber data. Setiap hasil diberi tanda sumbernya. Tanpa server, tanpa build.

| Tanda | File Excel asal |
|---|---|
| SSH 2026 | `export_excel_ssh_Kota_Malang_2026.xlsx` |
| SSH 2027 | `SHS_2027.xlsx` |
| SBU 2026 | `export_excel_sbu_Kota_Malang.xlsx` |
| ASB 2026 | `export_excel_asb_Kota_Malang.xlsx` |

## File
- `index.html` : seluruh tampilan dan logika pencarian
- `data.js` : data gabungan dari 4 file Excel (17.442 barang)
- `update_data.py` : membuat ulang `data.js` dari file Excel
- `.nojekyll` : agar GitHub Pages menyajikan file apa adanya

## Deploy ke GitHub Pages
1. Buat repository baru di GitHub, lalu unggah semua file ini ke branch `main`.
2. Buka Settings > Pages.
3. Pada Source pilih "Deploy from a branch", branch `main`, folder `/ (root)`, lalu Save.
4. Tunggu 1 sampai 2 menit. Alamatnya: `https://<username>.github.io/<nama-repo>/`

## Memperbarui data
Letakkan keempat file Excel di folder ini, lalu:
```
pip install pandas openpyxl
python update_data.py
```
Atau tentukan file sendiri: `python update_data.py --ssh2026 a.xlsx --ssh2027 b.xlsx --sbu c.xlsx --asb d.xlsx`

Lalu commit dan push `data.js`.

## Catatan data
- SSH 2027 (`SHS_2027.xlsx`) tidak memuat ID, kode barang, dan nama kelompok. Nama kelompok
  diambil dari file lain berdasarkan kode kelompok; baris yang kembar persis dibuang.
- Induk kode rekening (format SIPD) diambil dari `data.js` yang sudah ada. Untuk melengkapi,
  jalankan skrip dengan `--rekening daftar_rekening.xlsx` (kolom `Kode Rekening` dan `Uraian`).

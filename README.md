# Cari Barang SSH Kota Malang 2027

Website statis untuk mencari nama barang, kode rekening, dan harga satuan dari
Standar Satuan Harga (SSH) Kota Malang 2027. Tanpa server, tanpa build.

## File
- `index.html` : seluruh tampilan dan logika pencarian
- `data.js` : data dari Excel (6.934 barang)
- `update_data.py` : membuat ulang `data.js` dari file Excel baru
- `.nojekyll` : agar GitHub Pages menyajikan file apa adanya

## Deploy ke GitHub Pages
1. Buat repository baru di GitHub, lalu unggah semua file ini ke branch `main`.
2. Buka Settings > Pages.
3. Pada Source pilih "Deploy from a branch", branch `main`, folder `/ (root)`, lalu Save.
4. Tunggu 1 sampai 2 menit. Alamatnya: `https://<username>.github.io/<nama-repo>/`

## Memperbarui data
```
python update_data.py file_excel_baru.xlsx
```
Lalu commit dan push `data.js`.

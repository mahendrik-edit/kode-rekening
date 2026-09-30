"""Ubah file Excel SSH menjadi data.js untuk website.

Pemakaian:
    pip install pandas openpyxl
    python update_data.py nama_file.xlsx

Kolom yang dibutuhkan (baris pertama sebagai judul):
KODE KELOMPOK BARANG, URAIAN KELOMPOK BARANG, ID STANDAR HARGA, KODE BARANG,
URAIAN BARANG, SPESIFIKASI, SATUAN, HARGA SATUAN, KODE REKENING
"""
import json, re, sys
import pandas as pd

src = sys.argv[1] if len(sys.argv) > 1 else 'export_excel_ssh_Kota_Malang_2027_updated.xlsx'
df = pd.read_excel(src, dtype=str).fillna('')

def clean(v):
    v = re.sub(r'\s+', ' ', str(v)).strip()
    return '' if v in ('-', 'nan', 'NaN') else v

groups, gidx, units, uidx, accts, aidx, rows = [], {}, [], {}, [], {}, []
for _, r in df.iterrows():
    g = (clean(r['KODE KELOMPOK BARANG']), clean(r['URAIAN KELOMPOK BARANG']))
    if g not in gidx:
        gidx[g] = len(groups); groups.append(list(g))
    u = clean(r['SATUAN'])
    if u not in uidx:
        uidx[u] = len(units); units.append(u)
    a = clean(r['KODE REKENING'])
    if a:
        if a not in aidx:
            aidx[a] = len(accts); accts.append(a)
        ai = aidx[a]
    else:
        ai = -1
    rows.append([int(r['ID STANDAR HARGA']), clean(r['KODE BARANG']), clean(r['URAIAN BARANG']),
                 clean(r['SPESIFIKASI']), uidx[u], int(float(r['HARGA SATUAN'])), ai, gidx[g]])

data = {'groups': groups, 'units': units, 'accts': accts, 'rows': rows}
with open('data.js', 'w', encoding='utf-8') as f:
    f.write('window.SSH=' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';')
print(f'Selesai: {len(rows)} barang, {len(groups)} kelompok -> data.js')

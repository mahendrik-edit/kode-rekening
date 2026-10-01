"""Ubah file Excel menjadi data.js untuk website.

Pemakaian:
    pip install pandas openpyxl
    python update_data.py file_ssh.xlsx file_kode_rekening.xlsx

- file_ssh.xlsx          : export SSH (judul kolom di baris pertama):
  KODE KELOMPOK BARANG, URAIAN KELOMPOK BARANG, ID STANDAR HARGA, KODE BARANG,
  URAIAN BARANG, SPESIFIKASI, SATUAN, HARGA SATUAN, KODE REKENING
- file_kode_rekening.xlsx: daftar rekening SIPD, kolom "Kode Rekening" dan "Uraian"
  (opsional; tanpa file ini, induk kode rekening tidak ditampilkan)
"""
import json, re, sys
import pandas as pd

ssh_file = sys.argv[1] if len(sys.argv) > 1 else 'export_excel_ssh_Kota_Malang_2027_updated.xlsx'
rek_file = sys.argv[2] if len(sys.argv) > 2 else None

def clean(v):
    v = re.sub(r'\s+', ' ', str(v)).strip()
    return '' if v in ('-', 'nan', 'NaN') else v

def to_sipd(code):
    """5.1.02.01.001.00024 (format SSH) -> ['5','1','02','01','01','0024'] (format SIPD)."""
    p = code.strip().split('.')
    if len(p) == 6 and len(p[4]) == 3 and len(p[5]) == 5:
        return p[:4] + [p[4][1:], p[5][1:]]
    return p if len(p) == 6 else None

df = pd.read_excel(ssh_file, dtype=str).fillna('')
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

data = {'groups': groups, 'units': units, 'accts': accts, 'rows': rows, 'rek': {}}

if rek_file:
    rk = pd.read_excel(rek_file, dtype=str)
    rk = rk[rk['Kode Rekening'].notna()]
    sipd = {r['Kode Rekening'].strip().rstrip('.'): clean(r['Uraian']) for _, r in rk.iterrows()}
    unknown = set()
    for a in accts:
        for c in a.split(','):
            p = to_sipd(c)
            if not p:
                unknown.add(c.strip()); continue
            for n in range(1, 7):
                k = '.'.join(p[:n])
                if k in sipd:
                    data['rek'][k] = sipd[k]
                elif n == 6:
                    unknown.add(c.strip())
    print(f'Rekening induk: {len(data["rek"])} entri; {len(unknown)} kode SSH tidak ada di daftar rekening')

with open('data.js', 'w', encoding='utf-8') as f:
    f.write('window.SSH=' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';')
print(f'Selesai: {len(rows)} barang, {len(groups)} kelompok -> data.js')

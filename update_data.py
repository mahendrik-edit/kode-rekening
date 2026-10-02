"""Ubah 4 file Excel menjadi data.js untuk website.

Pemakaian (file Excel di folder yang sama):
    pip install pandas openpyxl
    python update_data.py

Atau tentukan sendiri:
    python update_data.py --ssh2026 a.xlsx --ssh2027 b.xlsx --sbu c.xlsx --asb d.xlsx [--rekening rek.xlsx]

Sumber data (setiap barang diberi tanda sumbernya):
  SSH 2026 : export SIPD, kolom KODE KELOMPOK BARANG ... KODE REKENING
  SSH 2027 : SHS_2027.xlsx, kolom Kode, Uraian, Spek, Satuan, Harga, Rekening 1
  SBU 2026 : export SIPD (format sama dengan SSH 2026)
  ASB 2026 : export SIPD (format sama dengan SSH 2026)

--rekening (opsional): daftar rekening SIPD dengan kolom "Kode Rekening" dan "Uraian".
Tanpa opsi ini, induk kode rekening diambil dari data.js yang sudah ada.
"""
import argparse, json, os, re
import pandas as pd

ap = argparse.ArgumentParser()
ap.add_argument('--ssh2026', default='export_excel_ssh_Kota_Malang_2026.xlsx')
ap.add_argument('--ssh2027', default='SHS_2027.xlsx')
ap.add_argument('--sbu', default='export_excel_sbu_Kota_Malang.xlsx')
ap.add_argument('--asb', default='export_excel_asb_Kota_Malang.xlsx')
ap.add_argument('--rekening', default=None)
ap.add_argument('--out', default='data.js')
a = ap.parse_args()

SRC = ['SSH 2026', 'SSH 2027', 'SBU 2026', 'ASB 2026']


def clean(v):
    v = re.sub(r'\s+', ' ', str(v)).strip()
    return '' if v in ('-', 'nan', 'NaN', 'None') else v


def to_sipd(code):
    """5.1.02.01.001.00024 (format SSH) -> ['5','1','02','01','01','0024'] (format SIPD)."""
    p = code.strip().split('.')
    if len(p) == 6 and len(p[4]) == 3 and len(p[5]) == 5:
        return p[:4] + [p[4][1:], p[5][1:]]
    return p if len(p) == 6 else None


def load_export(path):
    df = pd.read_excel(path, dtype=str).fillna('')
    return [dict(gk=clean(r['KODE KELOMPOK BARANG']), gn=clean(r['URAIAN KELOMPOK BARANG']),
                 id=int(r['ID STANDAR HARGA']), kb=clean(r['KODE BARANG']),
                 name=clean(r['URAIAN BARANG']), spec=clean(r['SPESIFIKASI']),
                 unit=clean(r['SATUAN']), price=round(float(r['HARGA SATUAN'])),
                 acct=clean(r['KODE REKENING'])) for _, r in df.iterrows()]


def load_shs(path):
    df = pd.read_excel(path, dtype=str).fillna('')
    seen, out, dup = set(), [], 0
    for _, r in df.iterrows():
        row = dict(gk=clean(r['Kode']), gn='', id=0, kb='', name=clean(r['Uraian']),
                   spec=clean(r['Spek']), unit=clean(r['Satuan']),
                   price=round(float(r['Harga'])), acct=clean(r['Rekening 1']))
        k = tuple(row.values())
        if k in seen:
            dup += 1
            continue
        seen.add(k)
        out.append(row)
    print(f'SSH 2027: {dup} baris kembar persis dibuang')
    return out


data_by_src = [load_export(a.ssh2026), load_shs(a.ssh2027), load_export(a.sbu), load_export(a.asb)]

# nama kelompok untuk SSH 2027 (file sumbernya hanya berisi kode kelompok)
gname = {}
for i in (0, 2, 3):
    for r in data_by_src[i]:
        gname.setdefault(r['gk'], r['gn'])
nogroup = set()
for r in data_by_src[1]:
    r['gn'] = gname.get(r['gk'], '')
    if not r['gn']:
        nogroup.add(r['gk'])
        r['gn'] = 'Kelompok ' + r['gk']
if nogroup:
    print(f'Peringatan: {len(nogroup)} kode kelompok SSH 2027 tidak punya nama:', sorted(nogroup))

groups, gidx, units, uidx, accts, aidx, rows = [], {}, [], {}, [], {}, []
for s, lst in enumerate(data_by_src):
    for r in lst:
        g = (r['gk'], r['gn'])
        if g not in gidx:
            gidx[g] = len(groups); groups.append(list(g))
        if r['unit'] not in uidx:
            uidx[r['unit']] = len(units); units.append(r['unit'])
        ai = -1
        if r['acct']:
            if r['acct'] not in aidx:
                aidx[r['acct']] = len(accts); accts.append(r['acct'])
            ai = aidx[r['acct']]
        name = r['name'] or r['spec'] or r['kb'] or '(tanpa nama)'
        rows.append([r['id'], r['kb'], name, r['spec'] if r['name'] else '', uidx[r['unit']],
                     r['price'], ai, gidx[g], s])

# induk kode rekening: dari daftar rekening baru, atau pakai yang sudah ada di data.js
sipd = {}
if a.rekening:
    rk = pd.read_excel(a.rekening, dtype=str)
    rk = rk[rk['Kode Rekening'].notna()]
    sipd = {r['Kode Rekening'].strip().rstrip('.'): clean(r['Uraian']) for _, r in rk.iterrows()}
elif os.path.exists(a.out):
    try:
        old = open(a.out, encoding='utf-8').read().strip()
        sipd = json.loads(old[old.index('=') + 1:].rstrip(';'))['rek']
    except Exception:
        sipd = {}
rek, unknown = {}, set()
for ac in accts:
    for c in ac.split(','):
        p = to_sipd(c)
        if not p:
            unknown.add(c.strip()); continue
        for n in range(1, 7):
            k = '.'.join(p[:n])
            if k in sipd:
                rek[k] = sipd[k]
            elif n == 6:
                unknown.add(c.strip())
print(f'Induk rekening: {len(rek)} entri; {len(unknown)} kode rekening belum ada uraiannya')

out = {'src': SRC, 'groups': groups, 'units': units, 'accts': accts, 'rows': rows, 'rek': rek}
with open(a.out, 'w', encoding='utf-8') as f:
    f.write('window.SSH=' + json.dumps(out, ensure_ascii=False, separators=(',', ':')) + ';')
for s, n in enumerate(SRC):
    print(f'  {n}: {sum(1 for r in rows if r[8] == s)} barang')
print(f'Selesai: {len(rows)} barang, {len(groups)} kelompok -> {a.out}')

/* Logika cek harga: baca RKA (PDF SIPD / Excel) lalu cocokkan dengan database SSH.
   Murni fungsi, tanpa DOM, agar bisa diuji di Node. */
(function (root) {
'use strict';
var COLOR = { hijau: 1, kuning: 1, merah: 1, biru: 1, pink: 1, oranye: 1, ungu: 1, abu: 1 };
var STOP = new Set('dan untuk dengan yang per jenis tipe type ukuran spesifikasi atau dari pada'.split(' '));

function tok(s) {
  s = String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  s = s.replace(/([0-9])([a-z])/g, '$1 $2').replace(/[^a-z0-9]+/g, ' '); /* 80gr -> 80 gr, tetapi A4 tetap utuh */
  var o = [], seen = {};
  s.trim().split(' ').forEach(function (w) {
    if (!w) return;
    if (w === 'gsm' || w === 'gram' || w === 'g') w = 'gr';
    if (STOP.has(w) || seen[w]) return;
    seen[w] = 1; o.push(w);
    if (COLOR[w] && !seen.warna) { seen.warna = 1; o.push('warna'); }
  });
  return o;
}
function money(s) { return parseFloat(String(s).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.')); }
function normUnit(u) { return String(u || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
/* kode rekening ke format SIPD 6 bagian: 5.1.02.01.001.00025 -> 5.1.02.01.01.0025 */
function rekKey(c) {
  var p = String(c || '').trim().split('.');
  if (p.length === 6 && p[4].length === 3 && p[5].length === 5) return p.slice(0, 4).concat(p[4].slice(1), p[5].slice(1)).join('.');
  return p.length === 6 ? p.join('.') : '';
}

/* ---------- PDF cetak RKA SIPD: pages = [[{x,y,s}]] ---------- */
function parsePdf(pages) {
  var out = [], rek = '', cur = { lines: [], a: null };
  function flush() {
    if (cur.a) {
      var L = cur.lines, k = -1, i;
      for (i = 0; i < L.length; i++) if (/^(Spesifikasi|Alamat)\s*:/.test(L[i])) { k = i; break; }
      var name = (k < 0 ? L : L.slice(0, k)).join(' ');
      var spec = k < 0 ? '' : L.slice(k).join(' ').replace(/^(Spesifikasi|Alamat)\s*:\s*/, '');
      if (/^-?$/.test(spec.trim())) spec = '';
      var a = cur.a;
      out.push({ name: name.replace(/\s+/g, ' ').trim(), spec: spec.replace(/\s+/g, ' ').trim(), unit: a.unit,
        price: a.price, total: a.total, qty: a.price > 0 ? Math.round(a.total / a.price * 1000) / 1000 : 0, rek: rek, ref: 'Hal ' + a.page });
    }
    cur = { lines: [], a: null };
  }
  pages.forEach(function (its, pi) {
    its = its.filter(function (i) { return i.s.trim(); });
    function find(t) { return its.filter(function (i) { return i.s.trim() === t; })[0]; }
    var K = find('Koefisien'), S = find('Satuan'), P = find('PPN');
    if (!K || !S || !P) return;
    var xK = K.x - 8, xS = S.x - 6, xP = P.x, yTop = K.y - 5;
    its = its.filter(function (i) { return i.y < yTop && !/^SIPD-RI/.test(i.s); }).sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    its.forEach(function (i) {
      var s = i.s.trim();
      if (i.x < 100 && /^5(\.\d+){5}$/.test(s)) { flush(); rek = s; return; }
      if (i.x < 165 || /^Sumber Dana/.test(s)) { flush(); return; }
      if (i.x < xK) { cur.lines.push(s); return; }
      if (/^\d+([.,]\d+)?\s*%$/.test(s) && i.x > xP - 20 && i.x < xP + 40) {
        if (cur.a) { var keep = cur.lines; flush(); cur.lines = keep; }
        var near = function (f) { return its.filter(function (j) { return Math.abs(j.y - i.y) <= 8 && f(j); })[0]; };
        var pr = near(function (j) { return j.x >= xS + 30 && j.x < xP - 5 && /^[\d.]+,\d{2}$/.test(j.s.trim()); });
        var tt = near(function (j) { return j.x >= xP + 40 && /^Rp\./.test(j.s.trim()); });
        var un = its.filter(function (j) { return j.x >= xS && j.x < xS + 30 && Math.abs(j.y - i.y) <= 14; })
          .sort(function (a, b) { return b.y - a.y; }).map(function (j) { return j.s.trim(); }).join(' ');
        cur.a = { price: pr ? money(pr.s) : 0, total: tt ? money(tt.s) : 0, unit: un, page: pi + 1 };
      }
    });
  });
  flush();
  return out.filter(function (r) { return r.name; });
}

/* ---------- Excel RKA: rows = array of array (satu sheet) ---------- */
function parseSheet(rows) {
  var hi = -1, c = {}, i, j;
  for (i = 0; i < Math.min(rows.length, 60) && hi < 0; i++)
    for (j = 0; j < (rows[i] || []).length; j++) if (/^uraian$/i.test(String(rows[i][j] || '').trim())) { hi = i; c.u = j; break; }
  if (hi < 0) return null;
  for (i = hi; i <= hi + 3; i++) (rows[i] || []).forEach(function (v, j) {
    v = String(v || '').toLowerCase();
    if (/volume|koefisien/.test(v) && c.v == null) c.v = j;
    else if (/^satuan/.test(v) && c.s == null) c.s = j;
    else if (/harga/.test(v) && c.h == null) c.h = j;
    else if (/^jumlah/.test(v) && c.t == null) c.t = j;
  });
  if (c.h == null) return null;
  var out = [], rek = '';
  var num = function (v) { var n = typeof v === 'number' ? v : parseFloat(String(v).replace(/\./g, '').replace(',', '.')); return isFinite(n) ? n : null; };
  for (i = hi + 1; i < rows.length; i++) {
    var r = rows[i] || [];
    var parts = [];
    for (j = 0; j < 6; j++) { var p = String(r[j] == null ? '' : r[j]).trim().replace(/\.0$/, ''); if (p) parts.push(p); }
    if (parts.length === 6) { rek = parts.join('.'); continue; }
    var name = String(r[c.u] == null ? '' : r[c.u]).trim();
    var pr = num(r[c.h + 1]); if (!(pr > 0)) pr = num(r[c.h]);
    if (!name || !(pr > 0) || parts.length) continue;
    var tot = c.t != null ? num(r[c.t]) : null, q = c.v != null ? num(r[c.v]) : null;
    if (q == null) q = tot != null && pr ? tot / pr : 0;
    out.push({ name: name, spec: '', unit: c.s != null ? String(r[c.s] || '').trim() : '', price: pr,
      total: tot != null ? tot : q * pr, qty: Math.round(q * 1000) / 1000, rek: rek, ref: 'Baris ' + (i + 1) });
  }
  return out;
}

/* ---------- pencocokan ---------- */
function buildIndex(D) {
  var post = {}, toks = [], reks = [], N = D.rows.length;
  D.rows.forEach(function (r, i) {
    var t = tok(r[2] + ' ' + r[3]); toks.push(t);
    t.forEach(function (w) { (post[w] || (post[w] = [])).push(i); });
    var rs = {}; (r[6] >= 0 ? D.accts[r[6]] : '').split(',').forEach(function (a) { var k = rekKey(a); if (k) rs[k] = 1; });
    reks.push(rs);
  });
  var w = function (t) { return Math.log(1 + N / ((post[t] || [1]).length)); };
  return { D: D, post: post, toks: toks, reks: reks, w: w };
}

/* allowed: array boolean per sumber; mengembalikan kandidat terbaik. Tahun terbaru (SSH 2027) diutamakan. */
function match(q, X, allowed) {
  var Q = tok(q.name + ' ' + q.spec), sumQ = 0, hit = {}, rk = rekKey(q.rek), qu = normUnit(q.unit);
  if (!Q.length) return [];
  Q.forEach(function (t) { sumQ += X.w(t); (X.post[t] || []).forEach(function (i) { hit[i] = (hit[i] || 0) + X.w(t); }); });
  var qn = Q.join(' '), res = [];
  Object.keys(hit).forEach(function (k) {
    var i = +k, r = X.D.rows[i];
    if (allowed && !allowed[r[8]]) return;
    var sumT = 0; X.toks[i].forEach(function (t) { sumT += X.w(t); });
    var sc = hit[k] / (0.7 * sumQ + 0.3 * sumT);
    /* score = kemiripan isi (0..1); rank = score + pemilih seri: rekening sama, satuan sama, tahun terbaru */
    var rank = sc + (rk && X.reks[i][rk] ? 0.03 : 0) + (qu && qu === normUnit(X.D.units[r[4]]) ? 0.02 : 0) + (r[8] === 1 ? 0.06 : 0);
    res.push({ i: i, score: sc, rank: rank });
  });
  res.sort(function (a, b) { return b.rank - a.rank || X.D.rows[b.i][5] - X.D.rows[a.i][5]; });
  return res.slice(0, 6);
}

var api = { tok: tok, rekKey: rekKey, parsePdf: parsePdf, parseSheet: parseSheet, buildIndex: buildIndex, match: match };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.CEK = api;
})(typeof window !== 'undefined' ? window : this);

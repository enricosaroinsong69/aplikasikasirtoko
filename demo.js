/* Mode coba: spreadsheet tiruan yang disimpan di perangkat ini.
   Berisi salinan Kode.gs dengan kode akses khusus mode coba (bukan kode asli toko).
   Aplikasi Kasir dan Gudang yang dibuka di peramban yang sama memakai data contoh yang sama. */
(function () {
  'use strict';
  const KUNCI = 'tb2_demo_sheets';
  window.BUAT_DEMO = function () {
    const TZ = (Intl.DateTimeFormat().resolvedOptions().timeZone) || 'Asia/Jakarta';
    function sel(v) {
      if (typeof v === 'string') {
        if (v.charAt(0) === "'") return v.slice(1);
        if (v === '') return '';
        if (/^-?\d+([.]\d+)?$/.test(v)) return Number(v);
        if (/^\d{4}-\d{2}-\d{2}$/.test(v)) { const p = v.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
        if (v.toUpperCase() === 'TRUE') return true;
        if (v.toUpperCase() === 'FALSE') return false;
      }
      return v;
    }
    class Sheet {
      constructor(name, rows) { this.name = name; this.rows = rows || []; }
      get(r, c) { const row = this.rows[r - 1]; return row && row[c - 1] !== undefined && row[c - 1] !== null ? row[c - 1] : ''; }
      set(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = sel(v); }
      getLastRow() { for (let i = this.rows.length; i >= 1; i--) if ((this.rows[i - 1] || []).some(x => x !== '' && x !== null && x !== undefined)) return i; return 0; }
      getLastColumn() { let m = 0; this.rows.forEach(r => r.forEach((x, i) => { if (x !== '' && x !== undefined && x !== null) m = Math.max(m, i + 1); })); return m; }
      getRange(a, b, c, d) {
        if (typeof a === 'string') {
          const m = /^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/.exec(a); const col = s => s.split('').reduce((t, ch) => t * 26 + ch.charCodeAt(0) - 64, 0);
          const r1 = m[2] ? +m[2] : 1, c1 = col(m[1]), c2 = m[3] ? col(m[3]) : c1, r2 = m[4] ? +m[4] : (m[3] ? Math.max(this.getLastRow(), 1) : r1);
          return new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
        }
        return new Range(this, a, b, c || 1, d || 1);
      }
      getName() { return this.name; }
      setFrozenRows() { }
    }
    class Range {
      constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
      getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j)); out.push(row); } return out; }
      getValue() { return this.sh.get(this.r, this.c); }
      setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return this; }
      setValue(v) { this.sh.set(this.r, this.c, v); return this; }
      getRow() { return this.r; }
      createTextFinder(text) {
        const rg = this; let entire = false;
        const hasil = () => { const out = []; for (let i = 0; i < rg.nr; i++) for (let j = 0; j < rg.nc; j++) { const v = String(rg.sh.get(rg.r + i, rg.c + j)); if (entire ? v === text : v.indexOf(text) >= 0) out.push(new Range(rg.sh, rg.r + i, rg.c + j, 1, 1)); } return out; };
        return { matchEntireCell(x) { entire = x; return this; }, findNext() { return hasil()[0] || null; }, findAll() { return hasil(); } };
      }
    }
    let sheets = {};
    const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)), getSpreadsheetTimeZone: () => TZ, getName: () => 'Mode coba' };
    function fmtDate(d, tz, f) {
      const p = {}; new Intl.DateTimeFormat('en-GB', { timeZone: tz || TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(d).forEach(x => { p[x.type] = x.value; });
      return f.replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute).replace('ss', p.second);
    }
    let nFoto = 0;
    const SpreadsheetApp = { getActiveSpreadsheet: () => ss, getActive: () => ss, flush() { }, getUi: () => ({ alert() { }, createMenu() { return { addItem() { return this; }, addToUi() { } }; } }) };
    const Utilities = { formatDate: fmtDate, parseDate: s => { const p = String(s).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }, base64Decode: () => [], newBlob: (b, t, n) => ({ name: n }) };
    const folder = { getId: () => 'demo', createFile() { nFoto++; return { getUrl: () => '#foto-contoh-' + nFoto }; } };
    const DriveApp = { getFolderById: () => folder, createFolder: () => folder };
    const LockService = { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) };
    const ContentService = { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } };
    const HtmlService = { createHtmlOutput: s => ({ s }) };
    const Session = { getScriptTimeZone: () => TZ };

    // Kode akses. Boleh diganti; setelah diganti buat versi penerapan baru.
// kasir  : aplikasi Kasir (nota jual, pembayaran bon, batal nota)
// gudang : aplikasi Gudang (pesanan, barang masuk, stok, data barang)
// pemilik: bisa dipakai di kedua aplikasi
const KODE_AKSES = {
  'demo-kasir': 'kasir',
  'demo-gudang': 'gudang',
  'demo-pemilik': 'pemilik'
};

const TAB = {
  barang: 'Barang', nota: 'Nota', rincian: 'Rincian Nota', bayar: 'Pembayaran',
  masuk: 'Barang Masuk', mutasi: 'Mutasi Stok', setel: 'Pengaturan'
};

// Kolom data yang ditulis aplikasi. Kolom "(otomatis)" berisi rumus dan tidak pernah ditulis.
const KOLOM = {
  barang: ['ID', 'Kode', 'Nama Barang', 'Kategori', 'Satuan', 'Harga Beli', 'Harga Jual', 'Stok Minimum', 'Stok Awal', 'Status', 'Diubah'],
  nota: ['ID', 'No Nota', 'Tanggal', 'Jam', 'Pelanggan', 'HP', 'Alamat', 'Subtotal', 'Diskon', 'Ongkos Kirim', 'Total', 'Laba Kotor',
    'Cara Bayar', 'Dibayar', 'Kembali', 'Jatuh Tempo', 'Pengambilan', 'Status Gudang', 'Petugas Gudang', 'Selesai Gudang', 'Kasir',
    'Catatan', 'Batal', 'Alasan Batal', 'Dicatat'],
  rincian: ['ID Nota', 'No Nota', 'Tanggal', 'ID Barang', 'Kode', 'Nama Barang', 'Banyaknya', 'Satuan', 'Harga', 'Jumlah', 'Harga Beli',
    'Laba Kotor', 'Status Gudang', 'Batal'],
  bayar: ['ID', 'Tanggal', 'ID Nota', 'No Nota', 'Pelanggan', 'Jumlah', 'Cara Bayar', 'Catatan', 'Kasir', 'Dicatat'],
  masuk: ['ID', 'No Penerimaan', 'Tanggal', 'Supplier', 'No Nota Supplier', 'Total', 'Cara Bayar', 'Dibayar', 'Jatuh Tempo',
    'Status Bayar', 'Petugas', 'Catatan', 'Foto Nota', 'Dicatat'],
  mutasi: ['ID', 'Tanggal', 'Jam', 'Jenis', 'Referensi', 'ID Barang', 'Nama Barang', 'Jumlah (+/-)', 'Satuan', 'Harga', 'Oleh', 'Catatan']
};

const IZIN = {
  ping: ['kasir', 'gudang'], muat: ['kasir', 'gudang'], kirim: ['kasir', 'gudang'], nota: ['kasir', 'gudang'],
  piutang: ['kasir'], ringkas: ['kasir', 'gudang'], pesanan: ['kasir', 'gudang'], kartu: ['kasir', 'gudang'], masuk: ['gudang']
};
const PERAN_DATA = { nota: 'kasir', bayar: 'kasir', batal: 'kasir', status: 'gudang', masuk: 'gudang', opname: 'gudang', barang: 'gudang' };
const STATUS_GUDANG = ['Menunggu', 'Disiapkan', 'Selesai'];

let SIMPAN = {};   // tembolok per permintaan: tabel & isi tabel

/* ===================== Pintu masuk ===================== */
function doPost(e) {
  let isi;
  try { isi = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return keluar({ ok: false, pesan: 'Permintaan tidak bisa dibaca.' }); }
  return keluar(jalankan(isi));
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.data) {
    try { return keluar(jalankan(JSON.parse(p.data))); }
    catch (err) { return keluar({ ok: false, pesan: 'Permintaan tidak bisa dibaca.' }); }
  }
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;font-size:16px">Penghubung aplikasi Kasir &amp; Gudang aktif. ' +
    'Salin alamat halaman ini (berakhiran /exec) ke aplikasi Kasir dan Gudang.</p>');
}

function keluar(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function jalankan(isi) {
  SIMPAN = {};
  const peran = KODE_AKSES[String(isi.kode || '').trim()];
  if (!peran) return { ok: false, kodeSalah: true, pesan: 'Kode akses salah. Periksa kode di pengaturan aplikasi.' };
  const izin = IZIN[isi.aksi];
  if (!izin || !AKSI[isi.aksi]) return { ok: false, pesan: 'Perintah tidak dikenal: ' + isi.aksi };
  if (peran !== 'pemilik' && izin.indexOf(peran) < 0) return { ok: false, pesan: 'Kode akses ' + peran + ' tidak boleh melakukan perintah ini.' };
  try {
    const hasil = AKSI[isi.aksi](isi, peran) || {};
    hasil.ok = true;
    hasil.peran = peran;
    hasil.waktu = new Date().toISOString();
    return hasil;
  } catch (err) {
    return { ok: false, pesan: String((err && err.message) || err) };
  }
}

/* ===================== Perintah ===================== */
const AKSI = {
  ping: function () { return { toko: pengaturan() }; },

  muat: function (isi, peran) {
    return {
      toko: pengaturan(),
      barang: daftarBarang(),
      pelanggan: peran === 'gudang' ? [] : daftarPelanggan(),
      supplier: peran === 'kasir' ? [] : daftarSupplier()
    };
  },

  kirim: function (isi, peran) {
    const paket = Array.isArray(isi.paket) ? isi.paket : [];
    const kunci = LockService.getScriptLock();
    kunci.waitLock(30000);
    const hasil = [];
    try {
      paket.forEach(function (op) {
        try {
          const f = OPS[op.tipe];
          if (!f) throw new Error('Jenis data tidak dikenal: ' + op.tipe);
          if (peran !== 'pemilik' && PERAN_DATA[op.tipe] !== peran) throw new Error('Kode akses ' + peran + ' tidak boleh mengirim data ' + op.tipe + '.');
          const r = f(op.data || {}, isi) || {};
          r.oid = op.oid; r.ok = true;
          hasil.push(r);
        } catch (err) {
          hasil.push({ oid: op.oid, ok: false, pesan: String((err && err.message) || err) });
        }
      });
      SpreadsheetApp.flush();
    } finally {
      kunci.releaseLock();
    }
    SIMPAN = {};
    return { hasil: hasil, barang: daftarBarang() };
  },

  nota: function (isi) {
    const q = String(isi.q || '').toLowerCase().trim();
    const dari = isi.dari || '0000-01-01', sampai = isi.sampai || '9999-12-31';
    let L = data('nota').filter(function (n) {
      if (isi.id) return n.ID === isi.id;
      const t = fTgl(n.Tanggal);
      if (t < dari || t > sampai) return false;
      if (!q) return true;
      return (String(n['No Nota']) + ' ' + n.Pelanggan + ' ' + n.HP).toLowerCase().indexOf(q) >= 0;
    });
    L.sort(urutTerbaru);
    L = L.slice(0, Number(isi.batas) || 300);
    return { nota: keObjNota(L) };
  },

  piutang: function () {
    const lun = pelunasan();
    const L = data('nota').filter(function (n) { return !benar(n.Batal) && sisaNota(n, lun) > 0; });
    L.sort(urutTerbaru);
    return { nota: keObjNota(L) };
  },

  pesanan: function () {
    const hari = fTgl(new Date());
    const aktif = data('nota').filter(function (n) { return !benar(n.Batal) && String(n['Status Gudang']) !== 'Selesai'; });
    const selesai = data('nota').filter(function (n) {
      return !benar(n.Batal) && String(n['Status Gudang']) === 'Selesai' && n['Selesai Gudang'] instanceof Date && fTgl(n['Selesai Gudang']) === hari;
    });
    aktif.sort(function (a, b) { return waktu(a.Dicatat) - waktu(b.Dicatat); });
    selesai.sort(function (a, b) { return waktu(b['Selesai Gudang']) - waktu(a['Selesai Gudang']); });
    return { aktif: keObjNota(aktif), selesai: keObjNota(selesai.slice(0, 60)) };
  },

  ringkas: function (isi) {
    const dari = isi.dari || fTgl(new Date()), sampai = isi.sampai || dari;
    const R = { penjualan: 0, nota: 0, laba: 0, batal: 0, masuk: {}, totalMasuk: 0, pelunasan: 0, piutang: 0, pesananAktif: 0, barangMasuk: 0 };
    const lun = pelunasan();
    data('nota').forEach(function (n) {
      const t = fTgl(n.Tanggal), batal = benar(n.Batal);
      if (!batal) { R.piutang += sisaNota(n, lun); if (String(n['Status Gudang']) !== 'Selesai') R.pesananAktif++; }
      if (t < dari || t > sampai) return;
      if (batal) { R.batal++; return; }
      R.penjualan += angka(n.Total); R.nota++; R.laba += angka(n['Laba Kotor']);
      const dib = angka(n.Dibayar);
      if (dib > 0) { const m = String(n['Cara Bayar']) === 'Bon' ? 'Uang muka bon' : String(n['Cara Bayar'] || 'Tunai'); R.masuk[m] = (R.masuk[m] || 0) + dib; R.totalMasuk += dib; }
    });
    const notaMap = {}; data('nota').forEach(function (n) { notaMap[n.ID] = n; });
    data('bayar').forEach(function (p) {
      const t = fTgl(p.Tanggal); if (t < dari || t > sampai) return;
      const n = notaMap[p['ID Nota']]; if (n && benar(n.Batal)) return;
      const m = String(p['Cara Bayar'] || 'Tunai'), j = angka(p.Jumlah);
      R.masuk[m] = (R.masuk[m] || 0) + j; R.totalMasuk += j; R.pelunasan += j;
    });
    data('masuk').forEach(function (m) { const t = fTgl(m.Tanggal); if (t >= dari && t <= sampai) R.barangMasuk += angka(m.Total); });
    return { ringkas: R };
  },

  kartu: function (isi) {
    const b = data('barang').filter(function (x) { return x.ID === isi.bid; })[0];
    if (!b) throw new Error('Barang tidak ditemukan.');
    let saldo = angka(b['Stok Awal']);
    const L = data('mutasi').filter(function (m) { return m['ID Barang'] === isi.bid; }).map(function (m) {
      return { tgl: fTgl(m.Tanggal), jam: String(m.Jam || ''), jenis: String(m.Jenis), ref: String(m.Referensi), qty: angka(m['Jumlah (+/-)']), oleh: String(m.Oleh || ''), ket: String(m.Catatan || ''), _r: m._r };
    });
    L.sort(function (a, b) { return a.tgl < b.tgl ? -1 : a.tgl > b.tgl ? 1 : a._r - b._r; });
    L.forEach(function (m) { saldo = bulat(saldo + m.qty); m.saldo = saldo; delete m._r; });
    const st = hitungStok();
    return { awal: angka(b['Stok Awal']), stok: st.stok[isi.bid] || 0, belum: st.belum[isi.bid] || 0, mutasi: L.reverse().slice(0, 80) };
  },

  masuk: function (isi) {
    const dari = isi.dari || '0000-01-01', sampai = isi.sampai || '9999-12-31';
    const per = {};
    data('mutasi').forEach(function (m) {
      if (String(m.Jenis) !== 'Masuk') return;
      const k = String(m.Referensi);
      (per[k] = per[k] || []).push({ bid: m['ID Barang'], nama: String(m['Nama Barang']), sat: String(m.Satuan || ''), qty: angka(m['Jumlah (+/-)']), harga: angka(m.Harga) });
    });
    let L = data('masuk').filter(function (m) { const t = fTgl(m.Tanggal); return t >= dari && t <= sampai; });
    L.sort(urutTerbaru);
    L = L.slice(0, 200).map(function (m) {
      return {
        id: m.ID, no: String(m['No Penerimaan']), tgl: fTgl(m.Tanggal), supplier: String(m.Supplier || ''), noSup: String(m['No Nota Supplier'] || ''),
        total: angka(m.Total), metode: String(m['Cara Bayar'] || ''), bayar: angka(m.Dibayar), jt: m['Jatuh Tempo'] ? fTgl(m['Jatuh Tempo']) : '',
        statusBayar: String(m['Status Bayar'] || ''), petugas: String(m.Petugas || ''), ket: String(m.Catatan || ''),
        foto: String(m['Foto Nota'] || '').split('\n').filter(String), items: per[String(m['No Penerimaan'])] || []
      };
    });
    return { masuk: L };
  }
};

/* ===================== Data yang dikirim aplikasi ===================== */
const OPS = {
  nota: function (d, isi) {
    if (!d.id || !d.no) throw new Error('Nota tanpa nomor.');
    if (cariBaris('nota', d.id)) return { id: d.id, no: d.no, ganda: true };
    const items = (d.items || []).filter(function (i) { return angka(i.qty) > 0 && String(i.nama || '').trim(); });
    if (!items.length) throw new Error('Nota ' + d.no + ' tidak berisi barang.');
    const bm = petaBarang();
    let sub = 0, laba = 0;
    const R = items.map(function (i) {
      const br = i.bid ? bm[i.bid] : null;
      const qty = bulat(angka(i.qty)), harga = Math.round(angka(i.harga));
      const jml = Math.round(qty * harga), hb = br ? angka(br['Harga Beli']) : 0;
      const lb = br ? jml - Math.round(qty * hb) : jml;
      sub += jml; laba += lb;
      return { i: i, br: br, qty: qty, harga: harga, jml: jml, hb: hb, lb: lb };
    });
    const diskon = Math.round(angka(d.diskon)), ongkir = Math.round(angka(d.ongkir));
    const total = Math.max(0, sub - diskon + ongkir);
    laba -= diskon;
    const metode = ['Tunai', 'Transfer', 'QRIS', 'Bon'].indexOf(d.metode) >= 0 ? d.metode : 'Tunai';
    let bayar, kembali = 0;
    if (metode === 'Tunai') { const terima = angka(d.terima) > 0 ? Math.round(angka(d.terima)) : total; bayar = Math.min(terima, total); kembali = Math.max(0, terima - total); }
    else if (metode === 'Bon') bayar = Math.min(Math.max(0, Math.round(angka(d.bayar))), total);
    else bayar = total;
    const tgl = keTanggal(d.tgl), kini = new Date();
    const status = d.ambil === 'Sudah dibawa' ? 'Selesai' : 'Menunggu';
    const kasir = d.kasir || isi.nama || '';
    tambah('nota', [{
      'ID': d.id, 'No Nota': T(d.no), 'Tanggal': tgl, 'Jam': T(d.jam || jamSekarang()), 'Pelanggan': T(d.nama), 'HP': T(d.hp), 'Alamat': T(d.alamat),
      'Subtotal': sub, 'Diskon': diskon, 'Ongkos Kirim': ongkir, 'Total': total, 'Laba Kotor': laba, 'Cara Bayar': metode, 'Dibayar': bayar,
      'Kembali': kembali, 'Jatuh Tempo': d.jt && bayar < total ? keTanggal(d.jt) : '', 'Pengambilan': T(d.ambil || 'Ambil di gudang'),
      'Status Gudang': status, 'Petugas Gudang': '', 'Selesai Gudang': status === 'Selesai' ? kini : '', 'Kasir': T(kasir),
      'Catatan': T(d.ket), 'Batal': false, 'Alasan Batal': '', 'Dicatat': kini
    }]);
    tambah('rincian', R.map(function (x) {
      return {
        'ID Nota': d.id, 'No Nota': T(d.no), 'Tanggal': tgl, 'ID Barang': x.br ? x.br.ID : '', 'Kode': T(x.br ? x.br.Kode : ''),
        'Nama Barang': T(x.i.nama), 'Banyaknya': x.qty, 'Satuan': T(x.i.sat || (x.br ? x.br.Satuan : '')), 'Harga': x.harga, 'Jumlah': x.jml,
        'Harga Beli': x.hb, 'Laba Kotor': x.lb, 'Status Gudang': status, 'Batal': false
      };
    }));
    const mut = R.filter(function (x) { return x.br; }).map(function (x, k) {
      return {
        'ID': d.id + '-j' + k, 'Tanggal': tgl, 'Jam': T(d.jam || jamSekarang()), 'Jenis': 'Jual', 'Referensi': T(d.no), 'ID Barang': x.br.ID,
        'Nama Barang': T(x.br['Nama Barang']), 'Jumlah (+/-)': -x.qty, 'Satuan': T(x.br.Satuan), 'Harga': x.harga, 'Oleh': T(kasir), 'Catatan': ''
      };
    });
    if (mut.length) tambah('mutasi', mut);
    return { id: d.id, no: d.no, total: total, laba: laba };
  },

  bayar: function (d, isi) {
    if (!d.id || !d.notaId) throw new Error('Data pembayaran tidak lengkap.');
    if (cariBaris('bayar', d.id)) return { id: d.id, ganda: true };
    const r = cariBaris('nota', d.notaId);
    if (!r) throw new Error('Nota untuk pembayaran ini tidak ditemukan.');
    const n = bacaBaris('nota', r);
    if (benar(n.Batal)) throw new Error('Nota ' + n['No Nota'] + ' sudah dibatalkan.');
    const sisa = sisaNota(n, pelunasan());
    if (sisa <= 0) throw new Error('Nota ' + n['No Nota'] + ' sudah lunas.');
    const jml = Math.min(Math.round(angka(d.jumlah)), sisa);
    if (jml <= 0) throw new Error('Jumlah pembayaran kosong.');
    tambah('bayar', [{
      'ID': d.id, 'Tanggal': keTanggal(d.tgl), 'ID Nota': d.notaId, 'No Nota': T(n['No Nota']), 'Pelanggan': T(n.Pelanggan), 'Jumlah': jml,
      'Cara Bayar': ['Tunai', 'Transfer', 'QRIS'].indexOf(d.metode) >= 0 ? d.metode : 'Tunai', 'Catatan': T(d.ket),
      'Kasir': T(d.kasir || isi.nama), 'Dicatat': new Date()
    }]);
    return { id: d.id, jumlah: jml, sisa: sisa - jml };
  },

  batal: function (d, isi) {
    const r = cariBaris('nota', d.notaId);
    if (!r) throw new Error('Nota yang dibatalkan tidak ditemukan.');
    const n = bacaBaris('nota', r);
    if (benar(n.Batal)) return { id: d.notaId, ganda: true };
    ubahSel('nota', r, { 'Batal': true, 'Alasan Batal': T(d.alasan) });
    cariSemua('rincian', d.notaId).forEach(function (rr) { ubahSel('rincian', rr, { 'Batal': true }); });
    const tgl = hariIniTgl(), jam = jamSekarang();
    const mut = data('rincian').filter(function (x) { return x['ID Nota'] === d.notaId && x['ID Barang']; }).map(function (x, k) {
      return {
        'ID': d.notaId + '-b' + k, 'Tanggal': tgl, 'Jam': T(jam), 'Jenis': 'Batal jual', 'Referensi': T(n['No Nota']), 'ID Barang': x['ID Barang'],
        'Nama Barang': T(x['Nama Barang']), 'Jumlah (+/-)': angka(x.Banyaknya), 'Satuan': T(x.Satuan), 'Harga': angka(x.Harga),
        'Oleh': T(d.oleh || isi.nama), 'Catatan': T(d.alasan)
      };
    });
    if (mut.length) tambah('mutasi', mut);
    return { id: d.notaId };
  },

  status: function (d, isi) {
    if (STATUS_GUDANG.indexOf(d.status) < 0) throw new Error('Status gudang tidak dikenal.');
    const r = cariBaris('nota', d.notaId);
    if (!r) throw new Error('Nota tidak ditemukan.');
    const n = bacaBaris('nota', r);
    if (benar(n.Batal)) throw new Error('Nota ' + n['No Nota'] + ' sudah dibatalkan kasir.');
    ubahSel('nota', r, { 'Status Gudang': d.status, 'Petugas Gudang': T(d.petugas || isi.nama), 'Selesai Gudang': d.status === 'Selesai' ? new Date() : '' });
    cariSemua('rincian', d.notaId).forEach(function (rr) { ubahSel('rincian', rr, { 'Status Gudang': d.status }); });
    return { id: d.notaId, status: d.status };
  },

  barang: function (d) {
    const nama = String(d.nama || '').trim();
    if (!d.id || !nama) throw new Error('Nama barang kosong.');
    const lain = data('barang').filter(function (b) { return b.ID !== d.id && String(b['Nama Barang']).trim().toLowerCase() === nama.toLowerCase() && String(b.Status) !== 'Arsip'; })[0];
    if (lain) throw new Error('Barang "' + nama + '" sudah ada.');
    const isi = {
      'Kode': T(d.kode), 'Nama Barang': T(nama), 'Kategori': T(d.kat || 'Lainnya'), 'Satuan': T(d.sat || 'pcs'),
      'Harga Beli': Math.round(angka(d.hb)), 'Harga Jual': Math.round(angka(d.hj)), 'Stok Minimum': angka(d.min),
      'Status': d.status === 'Arsip' ? 'Arsip' : 'Aktif', 'Diubah': new Date()
    };
    const r = cariBaris('barang', d.id);
    if (r) ubahSel('barang', r, isi);
    else { isi.ID = d.id; isi['Stok Awal'] = angka(d.awal); tambah('barang', [isi]); }
    return { id: d.id };
  },

  masuk: function (d, isi) {
    if (!d.id || !d.no) throw new Error('Penerimaan barang tanpa nomor.');
    if (cariBaris('masuk', d.id)) return { id: d.id, no: d.no, ganda: true };
    const bm = petaBarang();
    const items = (d.items || []).filter(function (i) { return angka(i.qty) > 0; });
    if (!items.length) throw new Error('Penerimaan ' + d.no + ' tidak berisi barang.');
    items.forEach(function (i) { if (!bm[i.bid]) throw new Error('Barang "' + i.nama + '" belum ada di daftar barang.'); });
    let total = 0;
    items.forEach(function (i) { total += Math.round(angka(i.qty) * Math.round(angka(i.harga))); });
    total += Math.round(angka(d.ongkir));
    const metode = ['Tunai', 'Transfer', 'Tempo'].indexOf(d.metode) >= 0 ? d.metode : 'Tunai';
    const bayar = metode === 'Tempo' ? Math.min(Math.max(0, Math.round(angka(d.bayar))), total) : total;
    const tgl = keTanggal(d.tgl), petugas = d.petugas || isi.nama || '';
    const foto = simpanFoto(d.foto, d.no);
    tambah('masuk', [{
      'ID': d.id, 'No Penerimaan': T(d.no), 'Tanggal': tgl, 'Supplier': T(d.supplier), 'No Nota Supplier': T(d.noSup), 'Total': total,
      'Cara Bayar': metode, 'Dibayar': bayar, 'Jatuh Tempo': d.jt && bayar < total ? keTanggal(d.jt) : '',
      'Status Bayar': bayar >= total ? 'Lunas' : 'Belum lunas', 'Petugas': T(petugas), 'Catatan': T(d.ket), 'Foto Nota': foto, 'Dicatat': new Date()
    }]);
    tambah('mutasi', items.map(function (i, k) {
      const br = bm[i.bid];
      return {
        'ID': d.id + '-m' + k, 'Tanggal': tgl, 'Jam': T(d.jam || jamSekarang()), 'Jenis': 'Masuk', 'Referensi': T(d.no), 'ID Barang': br.ID,
        'Nama Barang': T(br['Nama Barang']), 'Jumlah (+/-)': bulat(angka(i.qty)), 'Satuan': T(br.Satuan), 'Harga': Math.round(angka(i.harga)),
        'Oleh': T(petugas), 'Catatan': T(d.supplier)
      };
    }));
    items.forEach(function (i) {
      const br = bm[i.bid], ubah = { 'Diubah': new Date() };
      if (angka(i.harga) > 0) ubah['Harga Beli'] = Math.round(angka(i.harga));
      if (angka(i.hjBaru) > 0) ubah['Harga Jual'] = Math.round(angka(i.hjBaru));
      ubahSel('barang', br._r, ubah);
    });
    return { id: d.id, no: d.no, total: total, foto: foto };
  },

  opname: function (d, isi) {
    if (!d.id) throw new Error('Data hitung stok tanpa ID.');
    if (data('mutasi').some(function (m) { return String(m.ID).indexOf(d.id + '-') === 0; })) return { id: d.id, ganda: true };
    const bm = petaBarang(), st = hitungStok(), tgl = keTanggal(d.tgl), jam = jamSekarang(), petugas = d.petugas || isi.nama || '';
    const rows = [], hasil = [];
    (d.items || []).forEach(function (i, k) {
      const br = bm[i.bid]; if (!br || i.fisik === '' || i.fisik === null || i.fisik === undefined) return;
      const harus = bulat((st.stok[i.bid] || 0) + (st.belum[i.bid] || 0));
      const selisih = bulat(angka(i.fisik) - harus);
      hasil.push({ bid: i.bid, selisih: selisih });
      if (!selisih) return;
      rows.push({
        'ID': d.id + '-' + k, 'Tanggal': tgl, 'Jam': T(jam), 'Jenis': 'Penyesuaian', 'Referensi': T(d.no || 'Hitung stok'), 'ID Barang': br.ID,
        'Nama Barang': T(br['Nama Barang']), 'Jumlah (+/-)': selisih, 'Satuan': T(br.Satuan), 'Harga': angka(br['Harga Beli']),
        'Oleh': T(petugas), 'Catatan': T(i.ket || 'Hitung fisik: ' + angka(i.fisik))
      });
    });
    if (rows.length) tambah('mutasi', rows);
    return { id: d.id, berubah: rows.length, selisih: hasil };
  }
};

/* ===================== Menyusun data ===================== */
function daftarBarang() {
  const st = hitungStok();
  return data('barang').map(function (b) {
    return {
      id: b.ID, kode: String(b.Kode || ''), nama: String(b['Nama Barang'] || ''), kat: String(b.Kategori || ''), sat: String(b.Satuan || ''),
      hb: angka(b['Harga Beli']), hj: angka(b['Harga Jual']), min: angka(b['Stok Minimum']), awal: angka(b['Stok Awal']),
      status: String(b.Status || 'Aktif') === 'Arsip' ? 'Arsip' : 'Aktif', stok: st.stok[b.ID] || 0, belum: st.belum[b.ID] || 0
    };
  });
}

function hitungStok() {
  if (SIMPAN._stok) return SIMPAN._stok;
  const stok = {}, belum = {};
  data('barang').forEach(function (b) { stok[b.ID] = angka(b['Stok Awal']); });
  data('mutasi').forEach(function (m) { const id = m['ID Barang']; if (id in stok) stok[id] += angka(m['Jumlah (+/-)']); });
  data('rincian').forEach(function (r) {
    if (benar(r.Batal) || String(r['Status Gudang']) === 'Selesai' || !r['ID Barang']) return;
    belum[r['ID Barang']] = (belum[r['ID Barang']] || 0) + angka(r.Banyaknya);
  });
  Object.keys(stok).forEach(function (k) { stok[k] = bulat(stok[k]); });
  Object.keys(belum).forEach(function (k) { belum[k] = bulat(belum[k]); });
  SIMPAN._stok = { stok: stok, belum: belum };
  return SIMPAN._stok;
}

function daftarPelanggan() {
  const m = {};
  data('nota').slice().sort(function (a, b) { return waktu(a.Dicatat) - waktu(b.Dicatat); }).forEach(function (n) {
    const nama = String(n.Pelanggan || '').trim(); if (!nama) return;
    const k = nama.toLowerCase(), lama = m[k] || {};
    m[k] = { nama: nama, hp: String(n.HP || '') || lama.hp || '', alamat: String(n.Alamat || '') || lama.alamat || '' };
  });
  return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return a.nama.localeCompare(b.nama); }).slice(0, 1500);
}

function daftarSupplier() {
  const m = {};
  data('masuk').forEach(function (r) { const s = String(r.Supplier || '').trim(); if (s) m[s.toLowerCase()] = s; });
  return Object.keys(m).map(function (k) { return m[k]; }).sort();
}

function pelunasan() {
  if (SIMPAN._lunas) return SIMPAN._lunas;
  const m = {};
  data('bayar').forEach(function (p) { m[p['ID Nota']] = (m[p['ID Nota']] || 0) + angka(p.Jumlah); });
  SIMPAN._lunas = m;
  return m;
}

function sisaNota(n, lun) {
  if (benar(n.Batal)) return 0;
  return Math.max(0, Math.round(angka(n.Total) - angka(n.Dibayar) - (lun[n.ID] || 0)));
}

function keObjNota(L) {
  if (!L.length) return [];
  const ids = {}; L.forEach(function (n) { ids[n.ID] = []; });
  data('rincian').forEach(function (r) {
    if (ids[r['ID Nota']]) ids[r['ID Nota']].push({
      bid: r['ID Barang'] || null, kode: String(r.Kode || ''), nama: String(r['Nama Barang'] || ''), sat: String(r.Satuan || ''),
      qty: angka(r.Banyaknya), harga: angka(r.Harga)
    });
  });
  const lun = pelunasan();
  return L.map(function (n) {
    const batal = benar(n.Batal);
    return {
      id: n.ID, no: String(n['No Nota']), tgl: fTgl(n.Tanggal), jam: String(n.Jam || ''), nama: String(n.Pelanggan || ''), hp: String(n.HP || ''),
      alamat: String(n.Alamat || ''), items: ids[n.ID], sub: angka(n.Subtotal), diskon: angka(n.Diskon), ongkir: angka(n['Ongkos Kirim']),
      total: angka(n.Total), laba: angka(n['Laba Kotor']), metode: String(n['Cara Bayar'] || 'Tunai'), bayar: angka(n.Dibayar),
      kembali: angka(n.Kembali), jt: n['Jatuh Tempo'] ? fTgl(n['Jatuh Tempo']) : '', ambil: String(n.Pengambilan || ''),
      status: String(n['Status Gudang'] || 'Menunggu'), petugas: String(n['Petugas Gudang'] || ''),
      selesai: n['Selesai Gudang'] instanceof Date ? n['Selesai Gudang'].toISOString() : '', kasir: String(n.Kasir || ''),
      ket: String(n.Catatan || ''), batal: batal, alasan: String(n['Alasan Batal'] || ''), lunas: lun[n.ID] || 0, sisa: sisaNota(n, lun),
      dibuat: n.Dicatat instanceof Date ? n.Dicatat.getTime() : 0
    };
  });
}

function petaBarang() {
  const m = {};
  data('barang').forEach(function (b) { m[b.ID] = b; });
  return m;
}

function pengaturan() {
  const p = bacaPengaturan();
  return {
    nama: p['Nama toko'] || 'Toko Bangunan', alamat: p['Alamat'] || '', telp: p['Telepon / WA'] || '', kaki: p['Catatan kaki nota'] || '',
    pj: p['Awalan nota jual'] || 'NJ', pb: p['Awalan barang masuk'] || 'BM'
  };
}

function bacaPengaturan() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB.setel);
  const p = {};
  if (!sh || sh.getLastRow() < 2) return p;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) p[String(r[0]).trim()] = String(r[1] == null ? '' : r[1]).trim(); });
  return p;
}

function tulisPengaturan(label, nilai) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TAB.setel) || ss.insertSheet(TAB.setel);
  const n = Math.max(sh.getLastRow(), 1);
  const v = sh.getRange(1, 1, n, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === label) { sh.getRange(i + 1, 2).setValue(nilai); return; }
  sh.getRange(n + 1, 1, 1, 2).setValues([[label, nilai]]);
}

function simpanFoto(daftar, no) {
  if (!daftar || !daftar.length) return '';
  const folder = folderFoto();
  return daftar.map(function (u, i) {
    const m = /^data:([^;]+);base64,(.*)$/.exec(String(u));
    if (!m) return '';
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], no + '-' + (i + 1) + '.jpg');
    return folder.createFile(blob).getUrl();
  }).filter(String).join('\n');
}

function folderFoto() {
  const id = bacaPengaturan()['Folder foto nota'];
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  const f = DriveApp.createFolder('Foto Nota Barang Masuk - ' + SpreadsheetApp.getActiveSpreadsheet().getName());
  tulisPengaturan('Folder foto nota', f.getId());
  return f;
}

/* ===================== Akses tabel ===================== */
function tabel(nama) {
  const kunci = 't_' + nama;
  if (SIMPAN[kunci]) return SIMPAN[kunci];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB[nama]);
  if (!sh) {
    sh = ss.insertSheet(TAB[nama]);
    sh.getRange(1, 1, 1, KOLOM[nama].length).setValues([KOLOM[nama]]);
    sh.setFrozenRows(1);
  }
  const lebar = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lebar).getValues()[0].map(function (h) { return String(h).trim(); });
  const idx = {};
  head.forEach(function (h, i) { if (h && !idx[h]) idx[h] = i + 1; });
  KOLOM[nama].forEach(function (k) {
    if (!idx[k]) throw new Error('Kolom "' + k + '" tidak ditemukan di tab ' + TAB[nama] + '. Judul kolom jangan diganti.');
  });
  SIMPAN[kunci] = { sh: sh, head: head, idx: idx, nama: nama };
  return SIMPAN[kunci];
}

function barisTerakhir(sh) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const a = sh.getRange(1, 1, n, 1).getValues();
  for (let i = n - 1; i >= 1; i--) if (a[i][0] !== '' && a[i][0] !== null) return i + 1;
  return 1;
}

function data(nama) {
  const kunci = 'd_' + nama;
  if (SIMPAN[kunci]) return SIMPAN[kunci];
  const t = tabel(nama);
  const last = barisTerakhir(t.sh);
  const out = [];
  if (last >= 2) {
    const v = t.sh.getRange(2, 1, last - 1, t.head.length).getValues();
    for (let i = 0; i < v.length; i++) {
      const r = v[i];
      if (r[0] === '' || r[0] === null) continue;
      const o = { _r: i + 2 };
      KOLOM[nama].forEach(function (k) { o[k] = r[t.idx[k] - 1]; });
      out.push(o);
    }
  }
  SIMPAN[kunci] = out;
  return out;
}

function lupakan(nama) {
  delete SIMPAN['d_' + nama];
  if (nama === 'barang' || nama === 'mutasi' || nama === 'rincian') delete SIMPAN._stok;
  if (nama === 'bayar') delete SIMPAN._lunas;
}

// Menulis baris baru hanya ke kolom data (kolom rumus tidak disentuh).
function tambah(nama, objs) {
  if (!objs.length) return;
  const t = tabel(nama);
  const mulai = barisTerakhir(t.sh) + 1;
  const kol = KOLOM[nama].map(function (k) { return { k: k, c: t.idx[k] }; }).sort(function (a, b) { return a.c - b.c; });
  let run = [];
  const tulis = function () {
    if (!run.length) return;
    const nilai = objs.map(function (o) { return run.map(function (x) { const v = o[x.k]; return v === undefined || v === null ? '' : v; }); });
    t.sh.getRange(mulai, run[0].c, objs.length, run.length).setValues(nilai);
    run = [];
  };
  kol.forEach(function (x) { if (run.length && x.c !== run[run.length - 1].c + 1) tulis(); run.push(x); });
  tulis();
  lupakan(nama);
}

function ubahSel(nama, baris, obj) {
  const t = tabel(nama);
  Object.keys(obj).forEach(function (k) {
    if (obj[k] === undefined || !t.idx[k]) return;
    t.sh.getRange(baris, t.idx[k]).setValue(obj[k]);
  });
  lupakan(nama);
}

function cariBaris(nama, id) {
  const t = tabel(nama);
  const last = barisTerakhir(t.sh);
  if (last < 2 || !id) return 0;
  const f = t.sh.getRange(2, 1, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}

function cariSemua(nama, id) {
  const t = tabel(nama);
  const last = barisTerakhir(t.sh);
  if (last < 2 || !id) return [];
  return t.sh.getRange(2, 1, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findAll().map(function (r) { return r.getRow(); });
}

function bacaBaris(nama, baris) {
  const t = tabel(nama);
  const r = t.sh.getRange(baris, 1, 1, t.head.length).getValues()[0];
  const o = { _r: baris };
  KOLOM[nama].forEach(function (k) { o[k] = r[t.idx[k] - 1]; });
  return o;
}

/* ===================== Alat bantu ===================== */
function T(v) {
  // Tanda petik di depan menjaga isian tetap teks (no HP, kode, nomor nota) dan mencegah rumus disisipkan.
  const s = String(v == null ? '' : v).trim();
  return s ? "'" + s : '';
}
function angka(v) { if (typeof v === 'number') return isFinite(v) ? v : 0; const n = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) ? n : 0; }
function bulat(v) { return Math.round(v * 1000) / 1000; }
function benar(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function zona() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }
function keTanggal(s) { s = String(s || '').slice(0, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return hariIniTgl(); return Utilities.parseDate(s, zona(), 'yyyy-MM-dd'); }
function hariIniTgl() { return Utilities.parseDate(Utilities.formatDate(new Date(), zona(), 'yyyy-MM-dd'), zona(), 'yyyy-MM-dd'); }
function fTgl(v) { if (v instanceof Date) return Utilities.formatDate(v, zona(), 'yyyy-MM-dd'); return String(v || '').slice(0, 10); }
function jamSekarang() { return Utilities.formatDate(new Date(), zona(), 'HH.mm'); }
function waktu(v) { return v instanceof Date ? v.getTime() : 0; }
function urutTerbaru(a, b) { const x = fTgl(a.Tanggal), y = fTgl(b.Tanggal); if (x !== y) return x < y ? 1 : -1; return waktu(b.Dicatat) - waktu(a.Dicatat); }

/* ===================== Menu di spreadsheet ===================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Kasir & Gudang')
    .addItem('Periksa tab data', 'periksaTab')
    .addItem('Lihat kode akses', 'lihatKode')
    .addToUi();
}

function periksaTab() {
  SIMPAN = {};
  const pesan = [];
  Object.keys(KOLOM).forEach(function (n) {
    try { tabel(n); pesan.push('Tab ' + TAB[n] + ': siap'); } catch (e) { pesan.push('Tab ' + TAB[n] + ': ' + e.message); }
  });
  SpreadsheetApp.getUi().alert(pesan.join('\n'));
}

function lihatKode() {
  const baris = Object.keys(KODE_AKSES).map(function (k) { return KODE_AKSES[k] + ': ' + k; });
  SpreadsheetApp.getUi().alert('Kode akses aplikasi\n\n' + baris.join('\n'));
}


    const KEPALA = {
      'Barang': ['ID', 'Kode', 'Nama Barang', 'Kategori', 'Satuan', 'Harga Beli', 'Harga Jual', 'Stok Minimum', 'Stok Awal', 'Status', 'Diubah', 'Stok (otomatis)', 'Belum Diambil (otomatis)', 'Nilai Stok (otomatis)'],
      'Nota': ['ID', 'No Nota', 'Tanggal', 'Jam', 'Pelanggan', 'HP', 'Alamat', 'Subtotal', 'Diskon', 'Ongkos Kirim', 'Total', 'Laba Kotor', 'Cara Bayar', 'Dibayar', 'Kembali', 'Jatuh Tempo', 'Pengambilan', 'Status Gudang', 'Petugas Gudang', 'Selesai Gudang', 'Kasir', 'Catatan', 'Batal', 'Alasan Batal', 'Dicatat', 'Pelunasan (otomatis)', 'Sisa (otomatis)'],
      'Rincian Nota': ['ID Nota', 'No Nota', 'Tanggal', 'ID Barang', 'Kode', 'Nama Barang', 'Banyaknya', 'Satuan', 'Harga', 'Jumlah', 'Harga Beli', 'Laba Kotor', 'Status Gudang', 'Batal', 'Kunci Belum Diambil (otomatis)'],
      'Pembayaran': ['ID', 'Tanggal', 'ID Nota', 'No Nota', 'Pelanggan', 'Jumlah', 'Cara Bayar', 'Catatan', 'Kasir', 'Dicatat'],
      'Barang Masuk': ['ID', 'No Penerimaan', 'Tanggal', 'Supplier', 'No Nota Supplier', 'Total', 'Cara Bayar', 'Dibayar', 'Jatuh Tempo', 'Status Bayar', 'Petugas', 'Catatan', 'Foto Nota', 'Dicatat', 'Sisa Utang (otomatis)'],
      'Mutasi Stok': ['ID', 'Tanggal', 'Jam', 'Jenis', 'Referensi', 'ID Barang', 'Nama Barang', 'Jumlah (+/-)', 'Satuan', 'Harga', 'Oleh', 'Catatan']
    };
    function simpanSheets() {
      const o = {}; Object.keys(sheets).forEach(n => { o[n] = sheets[n].rows.map(r => r.map(v => v instanceof Date ? { $d: v.getTime() } : v)); });
      try { localStorage.setItem(KUNCI, JSON.stringify(o)); } catch (e) { }
    }
    function muatSheets() {
      let o = null; try { o = JSON.parse(localStorage.getItem(KUNCI) || 'null'); } catch (e) { }
      if (!o || !o['Nota']) return false;
      sheets = {}; Object.keys(o).forEach(n => { sheets[n] = new Sheet(n, o[n].map(r => r.map(v => v && typeof v === 'object' && '$d' in v ? new Date(v.$d) : v))); });
      return true;
    }
    function tglStr(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
    function jamStr(d) { return String(d.getHours()).padStart(2, '0') + '.' + String(d.getMinutes()).padStart(2, '0'); }
    function isiContoh() {
      sheets = {};
      Object.keys(KEPALA).forEach(n => { const s = ss.insertSheet(n); KEPALA[n].forEach((x, i) => s.set(1, i + 1, x)); });
      const p = ss.insertSheet('Pengaturan');
      [['Pengaturan', 'Isi', 'Keterangan'], ['Nama toko', 'TB Contoh Jaya'], ['Alamat', 'Jl. Contoh No. 1 (mode coba)'], ['Telepon / WA', "'0812 0000 0000"],
        ['Catatan kaki nota', 'Barang yang sudah dibeli tidak dapat ditukar atau dikembalikan.'], ['Awalan nota jual', 'NJ'], ['Awalan barang masuk', 'BM'], ['Folder foto nota', '']]
        .forEach((r, i) => r.forEach((x, j) => p.set(i + 1, j + 1, x)));
      const B = [
        ['b01', '8991001', 'Semen Tonasa PCC 50 kg', 'Semen & Pasir', 'sak', 62000, 68000, 20, 120],
        ['b02', '', 'Semen Tiga Roda 40 kg', 'Semen & Pasir', 'sak', 54000, 59000, 10, 45],
        ['b03', '', 'Pasir beton', 'Semen & Pasir', 'm³', 230000, 275000, 2, 9],
        ['b04', '', 'Besi beton 8 mm SNI', 'Besi & Baja', 'batang', 52000, 58000, 30, 160],
        ['b05', '', 'Besi beton 10 mm SNI', 'Besi & Baja', 'batang', 79000, 87000, 20, 70],
        ['b06', '', 'Kawat bendrat', 'Paku, Baut & Kawat', 'kg', 21000, 26000, 5, 18],
        ['b07', '', 'Paku 5 cm', 'Paku, Baut & Kawat', 'kg', 19000, 23000, 5, 24],
        ['b08', '', 'Triplek 9 mm', 'Kayu & Triplek', 'lembar', 98000, 115000, 10, 40],
        ['b09', '', 'Kayu meranti 5/7 4 m', 'Kayu & Triplek', 'batang', 64000, 75000, 10, 6],
        ['b10', '', 'Seng gelombang 0,3 mm 6 kaki', 'Atap & Plafon', 'lembar', 58000, 66000, 20, 85],
        ['b11', '', 'Gypsum 9 mm', 'Atap & Plafon', 'lembar', 66000, 76000, 10, 30],
        ['b12', '', 'Pipa PVC 3/4" AW', 'Pipa & Sanitasi', 'batang', 31000, 38000, 10, 55],
        ['b13', '', 'Pipa PVC 4" D', 'Pipa & Sanitasi', 'batang', 92000, 108000, 5, 14],
        ['b14', '', 'Cat tembok putih 5 kg', 'Cat & Pelapis', 'pail', 118000, 140000, 5, 16],
        ['b15', '', 'Keramik 40x40 putih', 'Keramik & Lantai', 'dus', 54000, 63000, 10, 48],
        ['b16', '', 'Kabel NYM 2x1,5 mm', 'Listrik', 'roll', 405000, 465000, 2, 5]
      ];
      const K = { kode: 'demo-pemilik', nama: 'Data contoh' };
      jalankan(Object.assign({ aksi: 'kirim', paket: B.map(b => ({ oid: b[0], tipe: 'barang', data: { id: b[0], kode: b[1], nama: b[2], kat: b[3], sat: b[4], hb: b[5], hj: b[6], min: b[7], awal: b[8] } })) }, K));
      const kini = new Date(), kemarin = new Date(kini.getTime() - 864e5), ymd = d => tglStr(d).slice(2).replace(/-/g, '');
      const tK = tglStr(kemarin), tH = tglStr(kini), jt = tglStr(new Date(kini.getTime() + 14 * 864e5));
      const j1 = new Date(kini.getTime() - 47 * 6e4), j2 = new Date(kini.getTime() - 12 * 6e4);
      const N = [
        { id: 'contoh-n1', no: 'NJ-' + ymd(kemarin) + '-B001', tgl: tK, jam: '09.20', nama: '', items: [{ bid: 'b01', nama: 'Semen Tonasa PCC 50 kg', sat: 'sak', qty: 5, harga: 68000 }, { bid: 'b07', nama: 'Paku 5 cm', sat: 'kg', qty: 2, harga: 23000 }], metode: 'Tunai', terima: 400000, ambil: 'Sudah dibawa', kasir: 'Rina' },
        { id: 'contoh-n2', no: 'NJ-' + ymd(kemarin) + '-B002', tgl: tK, jam: '14.05', nama: 'Pak Yohanis', hp: "0813 1111 2222", items: [{ bid: 'b04', nama: 'Besi beton 8 mm SNI', sat: 'batang', qty: 20, harga: 58000 }, { bid: 'b06', nama: 'Kawat bendrat', sat: 'kg', qty: 3, harga: 26000 }], metode: 'Bon', bayar: 500000, jt: jt, ambil: 'Ambil nanti', ket: 'Diambil besok pagi pakai pikap', kasir: 'Rina' },
        { id: 'contoh-n3', no: 'NJ-' + ymd(kini) + '-B001', tgl: tH, jam: jamStr(j1), nama: 'Ibu Maria Lumowa', hp: '0852 3333 4444', alamat: 'Perum Griya Asri Blok C-7', items: [{ bid: 'b01', nama: 'Semen Tonasa PCC 50 kg', sat: 'sak', qty: 10, harga: 68000 }, { bid: 'b03', nama: 'Pasir beton', sat: 'm³', qty: 1, harga: 275000 }, { bid: 'b15', nama: 'Keramik 40x40 putih', sat: 'dus', qty: 12, harga: 63000 }], ongkir: 50000, metode: 'Transfer', ambil: 'Diantar', kasir: 'Rina' },
        { id: 'contoh-n4', no: 'NJ-' + ymd(kini) + '-B002', tgl: tH, jam: jamStr(j2), nama: 'Tukang Anton', items: [{ bid: 'b10', nama: 'Seng gelombang 0,3 mm 6 kaki', sat: 'lembar', qty: 15, harga: 66000 }, { bid: 'b09', nama: 'Kayu meranti 5/7 4 m', sat: 'batang', qty: 4, harga: 75000 }], metode: 'Tunai', terima: 1300000, ambil: 'Ambil di gudang', kasir: 'Rina' }
      ];
      jalankan(Object.assign({ aksi: 'kirim', paket: N.map(n => ({ oid: n.id, tipe: 'nota', data: n })) }, K));
      jalankan(Object.assign({ aksi: 'kirim', paket: [{ oid: 'contoh-s4', tipe: 'status', data: { notaId: 'contoh-n4', status: 'Disiapkan', petugas: 'Budi' } }] }, K));
      jalankan(Object.assign({ aksi: 'kirim', paket: [{ oid: 'contoh-m1', tipe: 'masuk', data: { id: 'contoh-m1', no: 'BM-' + ymd(kemarin) + '-G001', tgl: tK, jam: '10.40', supplier: 'UD Sinar Baja', noSup: 'SB-2231', items: [{ bid: 'b04', nama: 'Besi beton 8 mm SNI', qty: 50, harga: 52000, hjBaru: 58000 }, { bid: 'b05', nama: 'Besi beton 10 mm SNI', qty: 20, harga: 79000, hjBaru: 87000 }], ongkir: 0, metode: 'Tempo', bayar: 1500000, jt: jt, petugas: 'Budi', ket: '', foto: [] } }] }, K));
      // waktu dicatat disesuaikan dengan jam nota supaya urutan pesanan wajar
      const sh = sheets['Nota'], h = sh.rows[0], iT = h.indexOf('Tanggal'), iJ = h.indexOf('Jam'), iD = h.indexOf('Dicatat'), iS = h.indexOf('Selesai Gudang');
      for (let r = 1; r < sh.rows.length; r++) {
        const row = sh.rows[r]; if (!(row[iT] instanceof Date)) continue; const j = String(row[iJ]).split('.');
        row[iD] = new Date(row[iT].getFullYear(), row[iT].getMonth(), row[iT].getDate(), +j[0] || 0, +j[1] || 0);
        if (row[iS] instanceof Date) row[iS] = new Date(row[iD].getTime() + 15 * 6e4);
      }
      simpanSheets();
    }
    return {
      panggil(body) {
        if (!muatSheets()) isiContoh();
        const out = JSON.parse(doPost({ postData: { contents: JSON.stringify(body) } }).s);
        if (body.aksi === 'kirim') simpanSheets();
        return out;
      },
      ulang() { try { localStorage.removeItem(KUNCI); } catch (e) { } }
    };
  };
})();

/**
 * ============================================================================
 *  PEMBUKUAN & DASHBOARD MONEV OPERASIONAL AGROMAN — Google Apps Script
 * ============================================================================
 *  Cara pakai:
 *   1. Buka spreadsheet "Monitoring Agroman" → Ekstensi → Apps Script.
 *   2. Hapus isi Code.gs, tempel seluruh isi file ini, klik Simpan.
 *   3. Jalankan fungsi `bangunSemua` sekali (beri izin saat diminta),
 *      atau muat ulang spreadsheet lalu pakai menu "Agroman" → "Rapikan & bangun dashboard".
 *
 *  Yang dibuat (sheet sumber TIDAK diubah):
 *   Daftar Akun, Jurnal Kas, Rekap Bulanan, Rekap Wilayah, Rekap Periode Ops,
 *   Penerima, Anggaran Seragam, Pengajuan vs ACC, Pengiriman PJ, Rekonsiliasi,
 *   Dashboard (KPI + grafik).
 *
 *  Menjalankan ulang aman: sheet hasil lama dihapus lalu dibuat ulang dari sumber.
 *  Semua rekap memakai rumus (SUMIFS dll.) yang membaca Jurnal Kas, jadi baris
 *  baru yang ditambahkan manual di Jurnal Kas ikut terhitung.
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// KONFIGURASI — sesuaikan bila nama sheet sumber berbeda
// ----------------------------------------------------------------------------
const CFG = {
  SRC_SALDO: 'Monitoring Saldo',
  SRC_PENGAJUAN: 'Monitoring Pengajuan',
  SRC_PENGIRIMAN: 'Pengiriman Berkas',
  SRC_ANGGARAN: 'Anggaran',            // dicari tanpa memperhatikan spasi di ujung nama
  PERIODE_AWAL: [2023, 12, 2],         // periode ops pertama di buku kas (tahun, bulan, I/II)
  TANGGAL_AWAL: [2023, 12, 21],        // tanggal transaksi pertama (acuan pengisian tanggal kosong)
  REKENING_KAS: 'Rekening Operasional BNI',
  WARNA_HEADER: '#1F3A5F',
  FORMAT_ANGKA: '#,##0;(#,##0);"-"',
  FORMAT_TANGGAL: 'dd/mm/yyyy',
  FONT: 'Arial'
};

const SHEET = {
  AKUN: 'Daftar Akun', JURNAL: 'Jurnal Kas', BULANAN: 'Rekap Bulanan', WILAYAH: 'Rekap Wilayah',
  PERIODE: 'Rekap Periode Ops', PENERIMA: 'Penerima', SERAGAM: 'Anggaran Seragam',
  PENGAJUAN: 'Pengajuan vs ACC', PJ: 'Pengiriman PJ', REKON: 'Rekonsiliasi', DASH: 'Dashboard'
};

const AKUN = [
  ['4-101', 'Pencairan Uang Muka (UM) Operasional', 'Penerimaan', 'Masuk', 'Dana UM/Ops Agroman dari kantor pusat'],
  ['4-102', 'Penerimaan Kasbon', 'Penerimaan', 'Masuk', 'Pinjaman sementara (kasbon) yang wajib dikembalikan'],
  ['4-103', 'Penerimaan Reimbursement', 'Penerimaan', 'Masuk', 'Penggantian biaya yang sudah dikeluarkan'],
  ['4-104', 'Pengembalian Sisa Dana dari Lapangan', 'Penerimaan', 'Masuk', 'Sisa dana operasional yang dikembalikan Agroman'],
  ['5-101', 'Biaya Operasional Agroman', 'Pengeluaran', 'Keluar', 'Transfer operasional per periode dua mingguan ke Agroman'],
  ['5-102', 'Insentif Petugas Lapang', 'Pengeluaran', 'Keluar', 'Insentif triwulanan petugas lapang'],
  ['5-103', 'Biaya Seragam & Atribut', 'Pengeluaran', 'Keluar', 'Pembuatan/pembelian seragam, bordir'],
  ['5-104', 'Biaya Pengiriman Dokumen', 'Pengeluaran', 'Keluar', 'Ongkos kirim berkas (PUTS/PUTK, PJ)'],
  ['5-105', 'Biaya Perjalanan & Akomodasi', 'Pengeluaran', 'Keluar', 'Tiket, penginapan, reimburse perjalanan'],
  ['5-106', 'Biaya Administrasi Bank', 'Pengeluaran', 'Keluar', 'Biaya admin bulanan & biaya transfer antarbank'],
  ['6-101', 'Pengembalian UM / Sisa Dana ke Kas Pusat', 'Pengembalian Dana', 'Keluar', 'Setoran balik sisa UM/ops ke kantor pusat'],
  ['6-102', 'Pelunasan Kasbon', 'Pengembalian Dana', 'Keluar', 'Pengembalian kasbon ke kas pusat']
];

const BULAN = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember'];
const BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const WIL_FIX = { 'JABAR 1': 'JAWA BARAT 1', 'All Area': 'SEMUA AREA', '-': '' };

// ----------------------------------------------------------------------------
// MENU
// ----------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Agroman')
    .addItem('Rapikan & bangun dashboard', 'bangunSemua')
    .addItem('Perbarui grafik dashboard saja', 'bangunDashboard')
    .addToUi();
}

function bangunSemua() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Membaca & merapikan buku kas…', 'Agroman', 10);
  const L = parseBukuKas_(ss);
  const side = parseLainnya_(ss);
  [SHEET.DASH, SHEET.REKON, SHEET.PJ, SHEET.PENGAJUAN, SHEET.SERAGAM, SHEET.PENERIMA, SHEET.PERIODE,
   SHEET.WILAYAH, SHEET.BULANAN, SHEET.JURNAL, SHEET.AKUN].forEach((n) => { const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s); });

  ss.toast('Menulis Jurnal Kas…', 'Agroman', 10);
  tulisAkun_(ss);
  const J = tulisJurnal_(ss, L);
  SpreadsheetApp.flush();
  ss.toast('Menyusun rekap…', 'Agroman', 10);
  tulisBulanan_(ss, L);
  tulisWilayah_(ss, L);
  tulisPeriode_(ss, L);
  tulisPenerima_(ss, L);
  tulisSeragam_(ss, L, side);
  tulisPengajuan_(ss, side.pengajuan);
  tulisPJ_(ss, side.pengiriman);
  tulisRekon_(ss, L, J);
  SpreadsheetApp.flush();
  bangunDashboard();
  ss.toast('Selesai. ' + L.recs.length + ' mutasi dicatat; saldo akhir cocok: ' +
    (Math.round(L.recs.reduce((s, x) => s + x.masuk - x.keluar, 0)) === Math.round(L.finalBal) ? 'YA' : 'TIDAK'), 'Agroman', 15);
}

// ----------------------------------------------------------------------------
// UTILITAS
// ----------------------------------------------------------------------------
function num_(x) {
  if (x === null || x === undefined || x === '' || x === '-') return 0;
  if (typeof x === 'number') return x;
  let s = String(x).replace(/Rp/gi, '').replace(/\s/g, '');
  if (s === '' || s === '-') return 0;
  s = s.replace(/\./g, '').replace(',', '.');
  const v = parseFloat(s);
  return isNaN(v) ? 0 : v;
}
function numOrNull_(x) {
  if (x === null || x === undefined || x === '' || x === '-') return null;
  const v = num_(x);
  return (typeof x === 'number' || /\d/.test(String(x))) ? v : null;
}
// Tanggal disimpan pukul 12:00 agar tidak bergeser sehari bila zona waktu script ≠ zona waktu spreadsheet.
let SS_TZ_ = null;
function ssTz_() { return SS_TZ_ || (SS_TZ_ = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()); }
function d_(y, m, d) { return new Date(y, m - 1, d, 12); }
function addDays_(dt, n) { return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate() + n, 12); }
function dayOnly_(dt) { const p = Utilities.formatDate(dt, ssTz_(), 'yyyy-M-d').split('-'); return d_(+p[0], +p[1], +p[2]); }
function validDate_(y, m, d) { const t = d_(y, m, d); return t.getFullYear() === y && t.getMonth() === m - 1 && t.getDate() === d ? t : null; }
function ensureRows_(sh, n) { if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows()); }
function fmtD_(dt) { return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
function parseTgl_(x) {
  if (x instanceof Date && !isNaN(x)) return { date: dayOnly_(x), why: null };
  if (x === null || x === undefined || x === '' || x === '-') return { date: null, why: 'kosong' };
  const s = String(x).trim();
  const m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (m) {
    const t = validDate_(+m[3], +m[2], +m[1]);
    return t ? { date: t, why: 'teks' } : { date: null, why: 'tidak valid: ' + s };
  }
  return { date: null, why: 'bukan tanggal: ' + s };
}
function perLabel_(p) { return p[0] + '-' + ('0' + p[1]).slice(-2) + '-' + (p[2] === 1 ? 'I' : 'II'); }
function perNama_(k) { const a = k.split('-'); return BULAN_ID[+a[1] - 1] + ' ' + a[2] + ' ' + a[0]; }
function cmpPer_(a, b) { return a[0] - b[0] || a[1] - b[1] || a[2] - b[2]; }
function findSheet_(ss, name) {
  return ss.getSheets().find((s) => s.getName().trim().toLowerCase() === name.trim().toLowerCase()) || null;
}
function J_(col) { return "'" + SHEET.JURNAL + "'!$" + col + '$5:$' + col; }
function header_(sh, row, heads, widths) {
  const r = sh.getRange(row, 1, 1, heads.length);
  r.setValues([heads]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(CFG.WARNA_HEADER)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(row, 36);
  if (widths) widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}
function judul_(sh, t, sub) {
  sh.getRange('A1').setValue(t).setFontSize(14).setFontWeight('bold').setFontColor(CFG.WARNA_HEADER);
  if (sub) sh.getRange('A2').setValue(sub).setFontStyle('italic').setFontSize(9).setFontColor('#595959');
}
function rapikan_(sh) { sh.getDataRange().setFontFamily(CFG.FONT); }
function border_(rng) { rng.setBorder(true, true, true, true, true, true, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID); }

// ----------------------------------------------------------------------------
// 1. PARSE BUKU KAS (Monitoring Saldo)
// ----------------------------------------------------------------------------
function parseBukuKas_(ss) {
  const sh = findSheet_(ss, CFG.SRC_SALDO);
  if (!sh) throw new Error('Sheet "' + CFG.SRC_SALDO + '" tidak ditemukan.');
  const V = sh.getRange(1, 1, sh.getLastRow(), 9).getValues();
  let hdr = V.findIndex((r) => String(r[1]).trim().toLowerCase() === 'keterangan');
  if (hdr < 0) throw new Error('Baris judul "Keterangan" tidak ditemukan di ' + CFG.SRC_SALDO);
  let start = hdr + 1;
  while (start < V.length && String(V[start][1]).trim().toLowerCase() === 'keterangan') start++;
  // lompati baris sub-judul kosong
  while (start < V.length && V[start].every((c) => c === '' || c === null)) start++;

  let cur = CFG.PERIODE_AWAL.slice();
  let lastValid = d_(CFG.TANGGAL_AWAL[0], CFG.TANGGAL_AWAL[1], CFG.TANGGAL_AWAL[2]);
  let prevGrand = 0;
  const recs = [], markers = [], zeros = [], issues = [];

  for (let i = start; i < V.length; i++) {
    const r = V[i];
    if (r.every((c) => c === '' || c === null)) continue;
    const row = i + 1;
    let ket = String(r[1] || '').replace(/\s+/g, ' ').trim();
    const kl = ket.toLowerCase();
    const d = num_(r[4]), c = num_(r[5]), g = num_(r[6]);
    let wil = String(r[2] || '').trim();
    const note = [];
    if (/agrosol/i.test(wil)) { note.push("Kolom Wilayah berisi '" + wil + "' (dipindah ke uraian)"); ket += ' (' + wil + ')'; wil = ''; }
    if (WIL_FIX.hasOwnProperty(wil)) wil = WIL_FIX[wil];

    // periode ops
    const m = ket.match(/^op(?:erasional|s)\s+(?:periode\s+)?([a-z]+)\s*(\d)/i);
    const isOps = !!m && BULAN.indexOf(m[1].toLowerCase()) >= 0;
    let per;
    if (isOps) {
      const mo = BULAN.indexOf(m[1].toLowerCase()) + 1, half = +m[2] === 1 ? 1 : 2;
      const ym = ket.match(/(20\d\d)/);
      let y;
      if (ym) y = +ym[1];
      else { y = cur[0]; if (mo < cur[1] - 6) y++; else if (mo > cur[1] + 6) y--; }
      const p = [y, mo, half];
      if (cmpPer_(p, cur) >= 0 || ym) cur = p;
      per = p;
    } else per = cur.slice();

    // baris penanda "Sisa Saldo"
    if (/^(sisa )?saldo/i.test(kl) && d === 0 && c === 0) { markers.push({ row: row, ket: ket, saldo: g }); prevGrand = g; continue; }
    if (Math.abs(prevGrand + d - c - g) > 1) issues.push({ row: row, detail: prevGrand + ' + ' + d + ' - ' + c + ' ≠ ' + g });
    prevGrand = g;
    if (d === 0 && c === 0) { zeros.push({ row: row, ket: ket, wil: wil, per: per }); continue; }

    // tanggal
    const pt = parseTgl_(r[3]);
    const ref = isOps ? d_(per[0], per[1], per[2] === 1 ? 1 : 16) : lastValid;
    const lo = isOps ? addDays_(ref, -20) : addDays_(lastValid, -45);
    const hi = isOps ? addDays_(ref, 50) : addDays_(lastValid, 80);
    let date;
    if (pt.date && pt.date >= lo && pt.date <= hi) {
      date = pt.date;
      if (pt.why === 'teks') note.push("Tanggal ditulis sebagai teks '" + r[3] + "' → dikonversi");
    } else {
      let fixed = null;
      if (pt.date) {
        for (const yy of [ref.getFullYear(), ref.getFullYear() - 1, ref.getFullYear() + 1]) {
          const cand = validDate_(yy, pt.date.getMonth() + 1, pt.date.getDate());
          if (cand && cand >= lo && cand <= hi) { fixed = cand; break; }
        }
        if (!fixed && pt.date.getDate() <= 12) {
          const cand = validDate_(pt.date.getFullYear(), pt.date.getDate(), pt.date.getMonth() + 1);
          if (cand && cand >= lo && cand <= hi) fixed = cand;
        }
      }
      if (fixed) {
        date = fixed; note.push('Tanggal asli ' + fmtD_(pt.date) + ' tidak sesuai periode → dikoreksi ' + fmtD_(fixed));
      } else {
        date = isOps ? (lastValid > ref ? lastValid : ref) : lastValid;
        const asli = pt.why === 'kosong' ? 'kosong' : "'" + r[3] + "'";
        note.push('Tanggal ' + asli + ' → diisi ' + fmtD_(date) + ' (tanggal transaksi sebelumnya/awal periode)');
      }
    }
    if (date <= addDays_(lastValid, 80) && date > lastValid) lastValid = date;

    // akun
    let code;
    if (/admin/.test(kl)) code = '5-106';
    else if (/insentif/.test(kl)) code = '5-102';
    else if (/seragam|bordir/.test(kl)) code = '5-103';
    else if (/kirim/.test(kl)) code = '5-104';
    else if (/tiket|penginapan/.test(kl)) code = '5-105';
    else if (/pengembalian/.test(kl) && c > 0) code = /kasbon/.test(kl) ? '6-102' : '6-101';
    else if (/kasbon/.test(kl) && d > 0) code = '4-102';
    else if (/reimburs/.test(kl) && d > 0) code = '4-103';
    else if (/pencairan/.test(kl) && d > 0) code = '4-101';
    else if (isOps || /kekurangan ops/.test(kl)) code = c > 0 ? '5-101' : '4-104';
    else code = d > 0 ? '4-101' : '5-101';
    if (d > 0 && c > 0) note.push('Debet & kredit terisi bersamaan');

    recs.push({
      row: row, tanggal: date, per: per, uraian: ket, akun: code, wilayah: wil || 'Kantor/Umum',
      penerima: String(r[7] || '').replace(/\s+/g, ' ').trim().replace(/^-+|-+$/g, '').trim(),
      rekening: String(r[8] || '').trim().replace(/^-+|-+$/g, '').trim(),
      masuk: d, keluar: c, saldoSumber: g, note: note, ops: isOps
    });
  }

  // cek wilayah vs rekening (rekening yang biasanya dipakai wilayah lain)
  const byRek = {};
  recs.forEach((x) => { if (x.ops && x.rekening) { byRek[x.rekening] = byRek[x.rekening] || {}; byRek[x.rekening][x.wilayah] = (byRek[x.rekening][x.wilayah] || 0) + 1; } });
  recs.forEach((x) => {
    if (!x.ops || !x.rekening) return;
    const cnt = byRek[x.rekening], tot = Object.values(cnt).reduce((a, b) => a + b, 0);
    const mode = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
    if (x.wilayah !== mode && cnt[x.wilayah] === 1 && tot >= 5) x.note.push('Perlu cek: rekening ini biasanya untuk wilayah ' + mode);
  });
  recs.forEach((x) => {
    x.catatan = x.note.join('; ');
    if (!x.note.length) x.status = 'OK';
    else if (x.note.some((n) => /^Perlu cek/.test(n))) x.status = 'Perlu cek';
    else if (x.note.every((n) => /teks|Wilayah berisi/.test(n))) x.status = 'Info';
    else x.status = 'Dikoreksi';
  });
  return { recs: recs, markers: markers, zeros: zeros, issues: issues, finalBal: prevGrand };
}

// ----------------------------------------------------------------------------
// 2. PARSE PENGAJUAN, PENGIRIMAN BERKAS, ANGGARAN
// ----------------------------------------------------------------------------
function groupsOf_(V) {
  const hr = V.findIndex((r) => r.some((c) => String(c).trim() === 'Nama Agroman'));
  if (hr < 0) return null;
  const h1 = V[hr], h2 = V[hr + 1], groups = [];
  let cur = null;
  for (let j = 4; j < h1.length; j++) {
    if (h1[j] !== '' && h1[j] !== null) {
      let name = h1[j] instanceof Date ? BULAN_ID[h1[j].getMonth()] : String(h1[j]).replace(/\s+/g, ' ').trim();
      cur = { name: name, cols: {} }; groups.push(cur);
    }
    if (cur && h2[j]) cur.cols[String(h2[j]).trim()] = j;
  }
  return { hr: hr, groups: groups };
}
function perKeys_(groups) {
  let cur = [CFG.PERIODE_AWAL[0], CFG.PERIODE_AWAL[1]];
  return groups.map((g) => {
    const m = g.name.match(/([A-Za-z]+)\s*(I{1,2}|\d)?/);
    const mo = BULAN.indexOf((m ? m[1] : '').toLowerCase()) + 1;
    const half = m && (m[2] === 'II' || m[2] === '2') ? 2 : 1;
    const ym = g.name.match(/(20\d\d)/);
    let y = ym ? +ym[1] : cur[0];
    if (!ym) { if (mo < cur[1] - 6) y++; }
    cur = [y, mo];
    return mo > 0 ? perLabel_([y, mo, half]) : g.name;
  });
}
function parseLainnya_(ss) {
  const out = { pengajuan: [], pengiriman: [], anggaran: null };
  let sh = findSheet_(ss, CFG.SRC_PENGAJUAN);
  if (sh) {
    const V = sh.getDataRange().getValues(), G = groupsOf_(V);
    if (G) {
      const keys = perKeys_(G.groups);
      G.groups.forEach((g, gi) => {
        for (let i = G.hr + 2; i < V.length; i++) {
          const r = V[i];
          if (typeof r[0] !== 'number' || !r[1]) continue;
          const val = (k) => g.cols[k] !== undefined ? numOrNull_(r[g.cols[k]]) : null;
          const v = { pengajuan: val('Pengajuan'), acc: val('ACC'), koreksi: val('Koreksi'), sisa: val('Sisa Sebelumnya') };
          if (v.pengajuan === null && v.acc === null && v.koreksi === null && v.sisa === null) continue;
          out.pengajuan.push(Object.assign({ periode: keys[gi], agroman: String(r[1]).trim(), area: String(r[2] || '').trim(), base: String(r[3] || '').trim() }, v));
        }
      });
    }
  }
  sh = findSheet_(ss, CFG.SRC_PENGIRIMAN);
  if (sh) {
    const V = sh.getDataRange().getValues(), G = groupsOf_(V);
    if (G) {
      const keys = perKeys_(G.groups);
      G.groups.forEach((g, gi) => {
        for (let i = G.hr + 2; i < V.length; i++) {
          const r = V[i];
          if (typeof r[0] !== 'number' || !r[1]) continue;
          const k = r[g.cols['Tgl Kirim']], t = r[g.cols['Tgl Terima']], n = r[g.cols['No. Resi']];
          if ((k === '' || k == null) && (t === '' || t == null) && (n === '' || n == null)) continue;
          let resi = String(n || '').trim(), ket = '';
          if (typeof k === 'string' && k.trim()) ket = k.trim();
          if (resi === '-') resi = '';
          if (resi && !/\d/.test(resi)) { ket = (ket + ' ' + resi).trim(); resi = ''; }
          let eksp = '';
          [['J&T', /J\s*&\s*T|JNT|^JD|JD\d|JO\d/i], ['JNE', /JNE|^\d{12,}|CSS/i], ['TIKI', /TIKI|^6600/i], ['POS', /POS|^P2/i], ['Lion', /Lion|LP/i], ['ID Express', /^ID/i]]
            .some(([e, re]) => { if (re.test(resi)) { eksp = e; return true; } return false; });
          out.pengiriman.push({
            periode: keys[gi], agroman: String(r[1]).trim(), area: String(r[2] || '').trim(),
            kirim: k instanceof Date ? dayOnly_(k) : '', terima: t instanceof Date ? dayOnly_(t) : '', ekspedisi: eksp, resi: resi, ket: ket
          });
        }
      });
    }
  }
  sh = findSheet_(ss, CFG.SRC_ANGGARAN);
  if (sh) {
    const V = sh.getDataRange().getValues();
    const r = V.find((x) => /anggaran/i.test(String(x[2])) && num_(x[5]) > 0);
    if (r) out.anggaran = { tanggal: r[1] instanceof Date ? r[1] : '', ket: String(r[2]).trim(), jumlah: num_(r[5]), banyak: r[3], satuan: r[4] };
  }
  return out;
}

// ----------------------------------------------------------------------------
// 3. SHEET: DAFTAR AKUN
// ----------------------------------------------------------------------------
function tulisAkun_(ss) {
  const sh = ss.insertSheet(SHEET.AKUN);
  judul_(sh, 'Daftar Akun (Bagan Akun Kas Operasional)');
  header_(sh, 3, ['Kode', 'Nama Akun', 'Kelompok', 'Arus', 'Keterangan'], [70, 280, 140, 70, 420]);
  sh.getRange(4, 1, AKUN.length, 5).setValues(AKUN);
  border_(sh.getRange(4, 1, AKUN.length, 5));
  sh.setFrozenRows(3); rapikan_(sh);
}

// ----------------------------------------------------------------------------
// 4. SHEET: JURNAL KAS
// ----------------------------------------------------------------------------
function tulisJurnal_(ss, L) {
  const sh = ss.insertSheet(SHEET.JURNAL);
  judul_(sh, 'Jurnal Kas / Buku Bank — ' + CFG.REKENING_KAS,
    'Saldo awal 0. Kolom biru = input manual; kolom hitam = rumus. Untuk menambah transaksi: isi baris baru di bawah, lalu salin rumus kolom B, D, E, I, J, P dari baris di atasnya.');
  const heads = ['No', 'No. Bukti', 'Tanggal', 'Tahun', 'Bulan', 'Periode Ops', 'Uraian', 'Kode Akun', 'Nama Akun', 'Kelompok', 'Wilayah', 'Penerima',
    'Rekening Tujuan', 'Masuk (Rp)', 'Keluar (Rp)', 'Saldo (Rp)', 'Saldo Sumber (Rp)', 'Selisih', 'Status Data', 'Catatan Koreksi', 'Baris Sumber'];
  header_(sh, 4, heads, [40, 110, 85, 50, 75, 85, 300, 65, 230, 120, 150, 180, 180, 105, 105, 115, 115, 60, 85, 420, 60]);
  const R0 = 5, n = L.recs.length, AK = "'" + SHEET.AKUN + "'!$A$4:$A$" + (3 + AKUN.length);
  const seq = {}, rows = [];
  L.recs.forEach((x, i) => {
    const r = R0 + i, typ = x.masuk > 0 ? 'BKM' : 'BKK';
    const yymm = Utilities.formatDate(x.tanggal, Session.getScriptTimeZone(), 'yyMM');
    seq[typ + yymm] = (seq[typ + yymm] || 0) + 1;
    rows.push([
      i + 1, typ + '-' + yymm + '-' + ('00' + seq[typ + yymm]).slice(-3), x.tanggal,
      '=YEAR(C' + r + ')', '=DATE(YEAR(C' + r + '),MONTH(C' + r + '),1)', perLabel_(x.per), x.uraian, x.akun,
      '=IFERROR(INDEX(\'' + SHEET.AKUN + '\'!$B$4:$B$' + (3 + AKUN.length) + ',MATCH(H' + r + ',' + AK + ',0)),"KODE TIDAK ADA")',
      '=IFERROR(INDEX(\'' + SHEET.AKUN + '\'!$C$4:$C$' + (3 + AKUN.length) + ',MATCH(H' + r + ',' + AK + ',0)),"")',
      x.wilayah, x.penerima, x.rekening, x.masuk || '', x.keluar || '',
      i === 0 ? '=N' + r + '-O' + r : '=P' + (r - 1) + '+N' + r + '-O' + r,
      x.saldoSumber, '=ROUND(P' + r + '-Q' + r + ',0)', x.status, x.catatan, x.row
    ]);
  });
  ensureRows_(sh, R0 + n + 500);
  sh.getRange(R0, 1, n, heads.length).setValues(rows);
  const JL = R0 + n - 1;
  sh.getRange(R0, 3, n, 1).setNumberFormat(CFG.FORMAT_TANGGAL);
  sh.getRange(R0, 5, n, 1).setNumberFormat('mmm yyyy');
  sh.getRange(R0, 14, n, 5).setNumberFormat(CFG.FORMAT_ANGKA);
  [3, 6, 7, 8, 11, 12, 13, 14, 15].forEach((c) => sh.getRange(R0, c, n, 1).setFontColor('#0000FF'));
  border_(sh.getRange(R0, 1, n, heads.length));
  // warna status
  const bg = L.recs.map((x) => {
    const c = x.status === 'Dikoreksi' || x.status === 'Perlu cek' ? '#FFF2CC' : x.status === 'Info' ? '#EAF1FB' : null;
    return [c];
  });
  sh.getRange(R0, 19, n, 1).setBackgrounds(bg);
  sh.getRange(R0, 3, n, 1).setBackgrounds(bg.map((b) => [b[0] === '#FFF2CC' ? b[0] : null]));
  // validasi kode akun
  sh.getRange(R0, 8, sh.getMaxRows() - R0 + 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEET.AKUN).getRange(4, 1, AKUN.length, 1), true).setAllowInvalid(false).build());
  // selisih ≠ 0 merah
  const rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($R' + R0 + '<>"",$R' + R0 + '<>0)')
    .setBackground('#F8CBAD').setRanges([sh.getRange(R0, 18, sh.getMaxRows() - R0 + 1, 1)]).build();
  sh.setConditionalFormatRules([rule]);
  sh.setFrozenRows(4); sh.setFrozenColumns(7);
  sh.getRange(4, 1, n + 1, heads.length).createFilter();
  rapikan_(sh);
  return { R0: R0, JL: JL };
}

// ----------------------------------------------------------------------------
// 5. SHEET: REKAP BULANAN (laporan arus kas)
// ----------------------------------------------------------------------------
function tulisBulanan_(ss, L) {
  const sh = ss.insertSheet(SHEET.BULANAN);
  judul_(sh, 'Laporan Arus Kas Bulanan', 'Semua angka = SUMIFS dari Jurnal Kas (kolom Bulan & Kode Akun).');
  const pin = AKUN.filter((a) => a[3] === 'Masuk'), pout = AKUN.filter((a) => a[3] === 'Keluar');
  const heads = ['Bulan', 'Saldo Awal'].concat(pin.map((a) => a[0] + '\n' + a[1]), ['Total Masuk'], pout.map((a) => a[0] + '\n' + a[1]),
    ['Total Keluar', 'Saldo Akhir', 'Biaya Ops Bersih\n(5-101 − 4-104)', 'Cek (harus 0)']);
  header_(sh, 4, heads); sh.setRowHeight(4, 64);
  sh.setColumnWidth(1, 85); for (let c = 2; c <= heads.length; c++) sh.setColumnWidth(c, 115);
  const ds = L.recs.map((x) => x.tanggal.getTime());
  const min = new Date(Math.min.apply(null, ds)), max = new Date(Math.max.apply(null, ds));
  const months = [];
  for (let y = min.getFullYear(), m = min.getMonth(); y < max.getFullYear() || (y === max.getFullYear() && m <= max.getMonth());) {
    months.push([y, m + 1]); m++; if (m === 12) { m = 0; y++; }
  }
  const cTM = 3 + pin.length, cK0 = cTM + 1, cTK = cK0 + pout.length, cSA = cTK + 1, cNET = cSA + 1, cCHK = cNET + 1;
  const col = (c) => columnLetter_(c);
  const rows = months.map((mo, i) => {
    const r = 5 + i, row = ['=DATE(' + mo[0] + ',' + mo[1] + ',1)', i === 0 ? 0 : '=' + col(cSA) + (r - 1)];
    pin.forEach((a) => row.push('=SUMIFS(' + J_('N') + ',' + J_('E') + ',$A' + r + ',' + J_('H') + ',"' + a[0] + '")'));
    row.push('=SUMIFS(' + J_('N') + ',' + J_('E') + ',$A' + r + ')');
    pout.forEach((a) => row.push('=SUMIFS(' + J_('O') + ',' + J_('E') + ',$A' + r + ',' + J_('H') + ',"' + a[0] + '")'));
    row.push('=SUMIFS(' + J_('O') + ',' + J_('E') + ',$A' + r + ')');
    row.push('=B' + r + '+' + col(cTM) + r + '-' + col(cTK) + r);
    row.push('=' + col(cK0) + r + '-' + col(cTM - 1) + r);
    row.push('=ROUND(SUM(C' + r + ':' + col(cTM - 1) + r + ')-' + col(cTM) + r + '+SUM(' + col(cK0) + r + ':' + col(cTK - 1) + r + ')-' + col(cTK) + r + ',0)');
    return row;
  });
  sh.getRange(5, 1, rows.length, heads.length).setValues(rows);
  const ML = 4 + rows.length, tr = ML + 1;
  const tot = ['TOTAL', ''];
  for (let c = 3; c <= cCHK; c++) tot.push(c === cSA ? '=' + col(cSA) + ML : '=SUM(' + col(c) + '5:' + col(c) + ML + ')');
  sh.getRange(tr, 1, 1, heads.length).setValues([tot]).setFontWeight('bold').setBackground('#F2F2F2');
  sh.getRange(5, 1, rows.length, 1).setNumberFormat('mmm yyyy');
  sh.getRange(5, 2, rows.length + 1, heads.length - 1).setNumberFormat(CFG.FORMAT_ANGKA);
  [cTM, cTK, cSA].forEach((c) => sh.getRange(5, c, rows.length, 1).setBackground('#DCE6F1').setFontWeight('bold'));
  border_(sh.getRange(5, 1, rows.length + 1, heads.length));
  sh.setFrozenRows(4); sh.setFrozenColumns(2); rapikan_(sh);
}
function columnLetter_(c) { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }

// ----------------------------------------------------------------------------
// 6. SHEET: REKAP WILAYAH
// ----------------------------------------------------------------------------
function tulisWilayah_(ss, L) {
  const sh = ss.insertSheet(SHEET.WILAYAH);
  judul_(sh, 'Biaya Operasional Bersih per Wilayah', 'Biaya bersih = transfer operasional (5-101) − sisa dana dikembalikan (4-104).');
  const wils = Array.from(new Set(L.recs.filter((x) => x.akun === '5-101' || x.akun === '4-104').map((x) => x.wilayah))).sort();
  const yrs = Array.from(new Set(L.recs.map((x) => x.tanggal.getFullYear()))).sort();
  const heads = ['Wilayah'].concat(yrs.map((y) => 'Bersih ' + y), ['Total Bersih', 'Jumlah Transfer', 'Rata-rata / Transfer', 'Porsi']);
  header_(sh, 4, heads); sh.setColumnWidth(1, 170); for (let c = 2; c <= heads.length; c++) sh.setColumnWidth(c, 115);
  const cT = 2 + yrs.length, WL = 4 + wils.length;
  const rows = wils.map((w, i) => {
    const r = 5 + i, row = [w];
    yrs.forEach((y) => row.push('=SUMIFS(' + J_('O') + ',' + J_('K') + ',$A' + r + ',' + J_('D') + ',' + y + ',' + J_('H') + ',"5-101")-SUMIFS(' +
      J_('N') + ',' + J_('K') + ',$A' + r + ',' + J_('D') + ',' + y + ',' + J_('H') + ',"4-104")'));
    row.push('=SUM(B' + r + ':' + columnLetter_(cT - 1) + r + ')');
    row.push('=COUNTIFS(' + J_('K') + ',$A' + r + ',' + J_('H') + ',"5-101")');
    row.push('=IFERROR(' + columnLetter_(cT) + r + '/' + columnLetter_(cT + 1) + r + ',0)');
    row.push('=IFERROR(' + columnLetter_(cT) + r + '/' + columnLetter_(cT) + '$' + (WL + 1) + ',0)');
    return row;
  });
  sh.getRange(5, 1, rows.length, heads.length).setValues(rows);
  const tot = ['TOTAL'];
  for (let c = 2; c <= cT + 1; c++) tot.push('=SUM(' + columnLetter_(c) + '5:' + columnLetter_(c) + WL + ')');
  tot.push('=IFERROR(' + columnLetter_(cT) + (WL + 1) + '/' + columnLetter_(cT + 1) + (WL + 1) + ',0)', '');
  sh.getRange(WL + 1, 1, 1, heads.length).setValues([tot]).setFontWeight('bold').setBackground('#F2F2F2');
  sh.getRange(5, 2, rows.length + 1, heads.length - 2).setNumberFormat(CFG.FORMAT_ANGKA);
  sh.getRange(5, cT + 1, rows.length + 1, 1).setNumberFormat('#,##0');
  sh.getRange(5, heads.length, rows.length, 1).setNumberFormat('0.0%');
  border_(sh.getRange(5, 1, rows.length + 1, heads.length));
  sh.setFrozenRows(4); rapikan_(sh);
}

// ----------------------------------------------------------------------------
// 7. SHEET: REKAP PERIODE OPS
// ----------------------------------------------------------------------------
function tulisPeriode_(ss, L) {
  const sh = ss.insertSheet(SHEET.PERIODE);
  judul_(sh, 'Transfer Operasional per Periode (Dua Mingguan)', 'Periode I = tgl 1–15, II = tgl 16–akhir bulan.');
  const pers = Array.from(new Set(L.recs.filter((x) => x.akun === '5-101').map((x) => perLabel_(x.per)))).sort();
  const zc = {}; L.zeros.forEach((z) => { const k = perLabel_(z.per); zc[k] = (zc[k] || 0) + 1; });
  header_(sh, 4, ['Periode', 'Nama Periode', 'Total Transfer Ops (Rp)', 'Sisa Dikembalikan (Rp)', 'Biaya Bersih (Rp)', 'Jumlah Transfer', 'Rata-rata / Transfer', 'Transfer Rp0'],
    [90, 150, 140, 140, 130, 100, 130, 100]);
  const rows = pers.map((p, i) => {
    const r = 5 + i;
    return [p, perNama_(p), '=SUMIFS(' + J_('O') + ',' + J_('F') + ',$A' + r + ',' + J_('H') + ',"5-101")',
      '=SUMIFS(' + J_('N') + ',' + J_('F') + ',$A' + r + ',' + J_('H') + ',"4-104")', '=C' + r + '-D' + r,
      '=COUNTIFS(' + J_('F') + ',$A' + r + ',' + J_('H') + ',"5-101")', '=IFERROR(C' + r + '/F' + r + ',0)', zc[p] || 0];
  });
  sh.getRange(5, 1, rows.length, 8).setValues(rows);
  const PL = 4 + rows.length;
  sh.getRange(PL + 1, 1, 1, 8).setValues([['TOTAL', '', '=SUM(C5:C' + PL + ')', '=SUM(D5:D' + PL + ')', '=SUM(E5:E' + PL + ')', '=SUM(F5:F' + PL + ')',
    '=IFERROR(C' + (PL + 1) + '/F' + (PL + 1) + ',0)', '=SUM(H5:H' + PL + ')']]).setFontWeight('bold').setBackground('#F2F2F2');
  sh.getRange(5, 3, rows.length + 1, 5).setNumberFormat(CFG.FORMAT_ANGKA);
  sh.getRange(5, 6, rows.length + 1, 1).setNumberFormat('#,##0');
  border_(sh.getRange(5, 1, rows.length + 1, 8));
  sh.setFrozenRows(4); rapikan_(sh);
}

// ----------------------------------------------------------------------------
// 8. SHEET: PENERIMA
// ----------------------------------------------------------------------------
function tulisPenerima_(ss, L) {
  const sh = ss.insertSheet(SHEET.PENERIMA);
  judul_(sh, 'Daftar Penerima Dana Operasional', 'Total diterima = SUMIFS (akun 5-101) per rekening. "Aktif" = menerima transfer dalam 4 periode terakhir.');
  header_(sh, 4, ['Penerima', 'Rekening', 'Wilayah (terbanyak)', 'Periode Pertama', 'Periode Terakhir', 'Jumlah Transfer', 'Total Diterima (Rp)', 'Status'],
    [200, 200, 160, 100, 100, 100, 140, 90]);
  const P = {};
  L.recs.forEach((x) => {
    if (x.akun !== '5-101' || !x.rekening) return;
    const p = P[x.rekening] = P[x.rekening] || { names: {}, wil: {}, pers: [] };
    p.names[x.penerima] = (p.names[x.penerima] || 0) + 1; p.wil[x.wilayah] = (p.wil[x.wilayah] || 0) + 1; p.pers.push(perLabel_(x.per));
  });
  const allPers = Array.from(new Set(L.recs.filter((x) => x.akun === '5-101').map((x) => perLabel_(x.per)))).sort();
  const cutoff = allPers[Math.max(0, allPers.length - 4)];
  const top = (o) => Object.keys(o).sort((a, b) => o[b] - o[a])[0];
  const list = Object.keys(P).sort((a, b) => P[b].pers.length - P[a].pers.length);
  const rows = list.map((rek, i) => {
    const p = P[rek], r = 5 + i, ps = p.pers.slice().sort();
    return [top(p.names) || '(nama kosong)', rek, top(p.wil), ps[0], ps[ps.length - 1],
      '=COUNTIFS(' + J_('M') + ',$B' + r + ',' + J_('H') + ',"5-101")', '=SUMIFS(' + J_('O') + ',' + J_('M') + ',$B' + r + ',' + J_('H') + ',"5-101")',
      ps[ps.length - 1] >= cutoff ? 'Aktif' : 'Tidak aktif'];
  });
  if (rows.length) {
    sh.getRange(5, 1, rows.length, 8).setValues(rows);
    sh.getRange(5, 7, rows.length, 1).setNumberFormat(CFG.FORMAT_ANGKA);
    border_(sh.getRange(5, 1, rows.length, 8));
    sh.getRange(4, 1, rows.length + 1, 8).createFilter();
  }
  sh.setFrozenRows(4); rapikan_(sh);
}

// ----------------------------------------------------------------------------
// 9. SHEET: ANGGARAN SERAGAM
// ----------------------------------------------------------------------------
function tulisSeragam_(ss, L, side) {
  const sh = ss.insertSheet(SHEET.SERAGAM);
  judul_(sh, 'Anggaran vs Realisasi Seragam Agroman', 'Anggaran dari sheet Anggaran; realisasi ditarik dari Jurnal Kas akun 5-103.');
  const a = side.anggaran;
  sh.getRange('A4:B7').setValues([
    [a ? a.ket + ' (' + (a.banyak || '') + ' ' + (a.satuan || '') + ')' : 'Anggaran (isi manual)', a ? a.jumlah : 0],
    ['Realisasi (Jurnal Kas, akun 5-103)', '=SUMIFS(' + J_('O') + ',' + J_('H') + ',"5-103")'],
    ['Sisa Anggaran', '=B4-B5'], ['Penyerapan', '=IFERROR(B5/B4,0)']]);
  sh.getRange('B4').setFontColor('#0000FF');
  sh.getRange('B4:B6').setNumberFormat(CFG.FORMAT_ANGKA); sh.getRange('B7').setNumberFormat('0.0%');
  border_(sh.getRange('A4:B7'));
  header_(sh, 9, ['Tanggal', 'Uraian', 'Penerima', 'Jumlah (Rp)'], [320, 280, 200, 120]);
  const sr = L.recs.filter((x) => x.akun === '5-103');
  if (sr.length) {
    sh.getRange(10, 1, sr.length, 4).setValues(sr.map((x) => [x.tanggal, x.uraian, x.penerima, x.keluar]));
    sh.getRange(10, 1, sr.length, 1).setNumberFormat(CFG.FORMAT_TANGGAL);
    sh.getRange(10, 4, sr.length, 1).setNumberFormat(CFG.FORMAT_ANGKA);
    border_(sh.getRange(10, 1, sr.length, 4));
  }
  rapikan_(sh);
}

// ----------------------------------------------------------------------------
// 10. SHEET: PENGAJUAN vs ACC
// ----------------------------------------------------------------------------
function tulisPengajuan_(ss, data) {
  const sh = ss.insertSheet(SHEET.PENGAJUAN);
  judul_(sh, 'Pengajuan Dana vs Persetujuan (ACC)', 'ACC = disetujui Base Area; Koreksi = dari PJ terkirim; Sisa Sebelumnya = sisa UM − PJ terkirim.');
  header_(sh, 4, ['Periode', 'Nama Agroman', 'Area', 'Base Area', 'Pengajuan (Rp)', 'ACC (Rp)', 'Koreksi (Rp)', 'Sisa Sebelumnya (Rp)', 'Selisih Pengajuan−ACC', '% ACC', 'Catatan'],
    [90, 190, 90, 100, 115, 115, 115, 130, 130, 70, 300]);
  const d = data.slice().sort((a, b) => a.periode < b.periode ? -1 : a.periode > b.periode ? 1 : a.agroman < b.agroman ? -1 : 1);
  if (!d.length) return rapikan_(sh);
  const rows = d.map((x, i) => {
    const r = 5 + i;
    let note = '';
    if (x.acc !== null && x.acc < 0) note = 'ACC bernilai negatif — kemungkinan dana dikurangi sisa sebelumnya';
    if (x.pengajuan && x.acc === null) note = 'ACC kosong di sumber';
    const v = (z) => z === null ? '' : z;
    return [x.periode, x.agroman, x.area, x.base, v(x.pengajuan), v(x.acc), v(x.koreksi), v(x.sisa),
      '=IF(OR(E' + r + '="",F' + r + '=""),"",E' + r + '-F' + r + ')', '=IF(OR(E' + r + '="",E' + r + '=0,F' + r + '=""),"",F' + r + '/E' + r + ')', note];
  });
  sh.getRange(5, 1, rows.length, 11).setValues(rows);
  sh.getRange(5, 5, rows.length, 5).setNumberFormat(CFG.FORMAT_ANGKA);
  sh.getRange(5, 10, rows.length, 1).setNumberFormat('0%');
  border_(sh.getRange(5, 1, rows.length, 11));
  sh.getRange(4, 1, rows.length + 1, 11).createFilter();
  sh.setFrozenRows(4); rapikan_(sh);
}

// ----------------------------------------------------------------------------
// 11. SHEET: PENGIRIMAN PJ
// ----------------------------------------------------------------------------
function tulisPJ_(ss, data) {
  const sh = ss.insertSheet(SHEET.PJ);
  judul_(sh, 'Monitoring Pengiriman Berkas Pertanggungjawaban (PJ)', 'Lama kirim = Tgl Terima − Tgl Kirim (hari).');
  header_(sh, 4, ['Periode', 'Nama Agroman', 'Area', 'Tgl Kirim', 'Tgl Terima', 'Ekspedisi', 'No. Resi', 'Lama (hari)', 'Status', 'Keterangan'],
    [90, 190, 90, 90, 90, 80, 170, 80, 130, 220]);
  const d = data.slice().sort((a, b) => a.periode < b.periode ? -1 : a.periode > b.periode ? 1 : a.agroman < b.agroman ? -1 : 1);
  if (!d.length) return rapikan_(sh);
  const rows = d.map((x, i) => {
    const r = 5 + i;
    return [x.periode, x.agroman, x.area, x.kirim, x.terima, x.ekspedisi, x.resi,
      '=IF(AND(D' + r + '<>"",E' + r + '<>""),E' + r + '-D' + r + ',"")',
      '=IF(E' + r + '<>"","Diterima",IF(D' + r + '<>"","Dalam pengiriman","Tidak kirim"))', x.ket];
  });
  sh.getRange(5, 1, rows.length, 10).setValues(rows);
  sh.getRange(5, 4, rows.length, 2).setNumberFormat(CFG.FORMAT_TANGGAL);
  sh.getRange(5, 8, rows.length, 1).setNumberFormat('0');
  border_(sh.getRange(5, 1, rows.length, 10));
  sh.getRange(4, 1, rows.length + 1, 10).createFilter();
  sh.setFrozenRows(4); rapikan_(sh);
}

// ----------------------------------------------------------------------------
// 12. SHEET: REKONSILIASI
// ----------------------------------------------------------------------------
function tulisRekon_(ss, L, J) {
  const sh = ss.insertSheet(SHEET.REKON);
  judul_(sh, 'Rekonsiliasi & Catatan Kualitas Data');
  header_(sh, 3, ['Pemeriksaan', 'Nilai'], [420, 140]);
  const JN = "'" + SHEET.JURNAL + "'";
  const bl = ss.getSheetByName(SHEET.BULANAN), blLast = bl.getLastRow() - 1, blSA = bl.getLastColumn() - 2;
  const items = [
    ['Saldo akhir sheet sumber (Monitoring Saldo, kolom Grand)', L.finalBal],
    ['Saldo akhir Jurnal Kas (rumus)', '=' + JN + '!P' + J.JL],
    ['Saldo akhir Rekap Bulanan', "='" + SHEET.BULANAN + "'!" + columnLetter_(blSA) + blLast],
    ['Selisih Jurnal vs Sumber (harus 0)', '=ROUND(B5-B4,0)'],
    ['Selisih Rekap vs Jurnal (harus 0)', '=ROUND(B6-B5,0)'],
    ['Jumlah baris Selisih ≠ 0 di Jurnal', '=COUNTIF(' + JN + '!R5:R' + J.JL + ',"<>0")'],
    ['Jumlah mutasi dicatat', '=COUNTA(' + JN + '!B5:B' + J.JL + ')'],
    ['Baris "Sisa Saldo" sumber (tidak dicatat sebagai mutasi)', L.markers.length],
    ['Baris transfer Rp0 (tidak dicatat sebagai mutasi)', L.zeros.length],
    ['Mutasi dengan tanggal dikoreksi/diisi', '=COUNTIF(' + JN + '!S5:S' + J.JL + ',"Dikoreksi")'],
    ['Mutasi perlu dicek manual', '=COUNTIF(' + JN + '!S5:S' + J.JL + ',"Perlu cek")'],
    ['Saldo berjalan sumber yang tidak konsisten', L.issues.length]
  ];
  sh.getRange(4, 1, items.length, 2).setValues(items).setNumberFormat(CFG.FORMAT_ANGKA);
  sh.getRange(4, 1, items.length, 1).setNumberFormat('@');
  border_(sh.getRange(4, 1, items.length, 2));
  const r0 = 6 + items.length;
  sh.getRange(r0, 1).setValue('Daftar mutasi yang dikoreksi / perlu dicek').setFontWeight('bold');
  header_(sh, r0 + 1, ['Baris Sumber', 'Tanggal (bersih)', 'Uraian', 'Wilayah', 'Jumlah (Rp)', 'Status', 'Catatan']);
  [90, 100, 280, 150, 110, 90, 520].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setColumnWidth(1, 420);
  const fl = L.recs.filter((x) => x.status === 'Dikoreksi' || x.status === 'Perlu cek');
  if (fl.length) {
    sh.getRange(r0 + 2, 1, fl.length, 7).setValues(fl.map((x) => [x.row, x.tanggal, x.uraian, x.wilayah, x.masuk || x.keluar, x.status, x.catatan]));
    sh.getRange(r0 + 2, 2, fl.length, 1).setNumberFormat(CFG.FORMAT_TANGGAL);
    sh.getRange(r0 + 2, 5, fl.length, 1).setNumberFormat(CFG.FORMAT_ANGKA);
    border_(sh.getRange(r0 + 2, 1, fl.length, 7));
  }
  rapikan_(sh);
}

// ----------------------------------------------------------------------------
// 13. DASHBOARD (KPI + grafik)
// ----------------------------------------------------------------------------
function bangunDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET.JURNAL)) throw new Error('Jalankan "Rapikan & bangun dashboard" dulu.');
  let sh = ss.getSheetByName(SHEET.DASH);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET.DASH, 0);
  sh.setHiddenGridlines(true);
  sh.setColumnWidths(1, 12, 120);

  // ---- judul & ringkasan
  sh.getRange('A1').setValue('Monev Operasional Agroman').setFontSize(18).setFontWeight('bold').setFontColor(CFG.WARNA_HEADER);
  sh.getRange('A2:L2').merge().setWrap(true); sh.setRowHeight(2, 36);

  // ---- helper (kolom N–O): bulan terakhir & perubahan
  sh.getRange('N3:O3').setValues([['Bantu', '']]).setFontColor('#999999');
  sh.getRange('N4').setFormula('=MAX(' + J_('E') + ')').setNumberFormat('mmm yyyy');
  sh.getRange('O4').setFormula('=SUMIFS(' + J_('O') + ',' + J_('D') + ',YEAR(N4)-1,' + J_('E') + ',"<="&EDATE(N4,-12),' + J_('H') + ',"5-101")-SUMIFS(' + J_('N') + ',' + J_('D') + ',YEAR(N4)-1,' + J_('E') + ',"<="&EDATE(N4,-12),' + J_('H') + ',"4-104")')
    .setNumberFormat(CFG.FORMAT_ANGKA);
  sh.getRange('N5').setFormula('=IFERROR(G5/O4-1,0)').setNumberFormat('0.0%');

  // ---- KPI
  const kpiL = ['Saldo akhir rekening', 'Total penerimaan', 'Total pengeluaran', 'Biaya ops bersih tahun berjalan', 'Mutasi tercatat', 'Perlu dicek / dikoreksi'];
  const kpiF = [
    '=INDEX(' + J_('P') + ',COUNTA(' + J_('B') + '))',
    '=SUM(' + J_('N') + ')', '=SUM(' + J_('O') + ')',
    '=SUMIFS(' + J_('O') + ',' + J_('D') + ',YEAR(N4),' + J_('H') + ',"5-101")-SUMIFS(' + J_('N') + ',' + J_('D') + ',YEAR(N4),' + J_('H') + ',"4-104")',
    '=COUNTA(' + J_('B') + ')',
    '=COUNTIF(' + J_('S') + ',"Perlu cek")&" / "&COUNTIF(' + J_('S') + ',"Dikoreksi")'
  ];
  const cols = [1, 3, 5, 7, 9, 11];
  // KPI: label baris 4, nilai baris 5, masing-masing 2 kolom gabungan (saldo di A5, biaya ops di G5)
  kpiL.forEach((l, i) => {
    const c = cols[i];
    sh.getRange(4, c, 1, 2).merge().setValue(l).setFontColor('#595959').setFontSize(9);
    sh.getRange(5, c, 1, 2).merge().setFormula(kpiF[i]).setFontSize(16).setFontWeight('bold').setFontColor(CFG.WARNA_HEADER);
    sh.getRange(4, c, 2, 2).setBackground('#F3F6FA').setBorder(true, true, true, true, false, false, '#D0D7E2', SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange(5, 1, 1, 8).setNumberFormat('"Rp"#,##0');
  sh.getRange(5, 9).setNumberFormat('#,##0');
  // kalimat ringkasan dibaca dari KPI (A5 = saldo, G5 = biaya ops tahun berjalan)
  sh.getRange('A2').setFormula('="Per "&TEXT(N4,"mmmm yyyy")&", saldo rekening operasional Rp"&TEXT(A5,"#,##0")&". Biaya operasional bersih ke lapangan Jan–"&TEXT(N4,"mmm yyyy")&" Rp"&TEXT(G5,"#,##0")&", "&IF(N5>=0,"naik ","turun ")&TEXT(ABS(N5),"0.0%")&" dibanding periode sama tahun lalu."');

  // ---- data grafik (kolom N dst.)
  const D0 = 8;
  sh.getRange(D0, 14).setValue('DATA GRAFIK (jangan dihapus)').setFontWeight('bold').setFontColor('#999999');
  // a) komposisi pengeluaran
  const out = AKUN.filter((a) => a[3] === 'Keluar');
  sh.getRange(D0 + 1, 14, 1, 2).setValues([['Akun', 'Total Keluar']]).setFontWeight('bold');
  sh.getRange(D0 + 2, 14, out.length, 2).setValues(out.map((a, i) => [a[0] + ' ' + a[1], '=SUMIFS(' + J_('O') + ',' + J_('H') + ',"' + a[0] + '")']));
  sh.getRange(D0 + 2, 15, out.length, 1).setNumberFormat(CFG.FORMAT_ANGKA);
  const komp = sh.getRange(D0 + 1, 14, out.length + 1, 2);
  // b) wilayah untuk tahun terpilih
  const wl = ss.getSheetByName(SHEET.WILAYAH), wN = wl.getLastRow() - 5; // baris data wilayah
  const yrHead = wl.getRange(4, 2, 1, wl.getLastColumn() - 5).getValues()[0];
  const W0 = D0 + out.length + 4;
  sh.getRange(W0 - 1, 14).setValue('Tahun wilayah:').setFontWeight('bold');
  const pil = ['Semua tahun'].concat(yrHead.map((h) => String(h).replace('Bersih ', '')));
  sh.getRange(W0 - 1, 15).setValue(pil[pil.length - 1]).setBackground('#FFF2CC')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(pil, true).build());
  sh.getRange(W0, 14, 1, 2).setValues([['Wilayah', 'Biaya Bersih']]).setFontWeight('bold');
  const WS = "'" + SHEET.WILAYAH + "'";
  const wRows = [];
  for (let i = 0; i < wN; i++) {
    const r = 5 + i;
    wRows.push(['=' + WS + '!A' + r,
      '=IF($O$' + (W0 - 1) + '="Semua tahun",' + WS + '!' + columnLetter_(2 + yrHead.length) + r + ',IFERROR(INDEX(' + WS + '!B' + r + ':' + columnLetter_(1 + yrHead.length) + r +
      ',MATCH("Bersih "&$O$' + (W0 - 1) + ',' + WS + '!$B$4:$' + columnLetter_(1 + yrHead.length) + '$4,0)),0))']);
  }
  sh.getRange(W0 + 1, 14, wN, 2).setValues(wRows);
  sh.getRange(W0 + 1, 15, wN, 1).setNumberFormat(CFG.FORMAT_ANGKA);
  const wilR = sh.getRange(W0, 14, wN + 1, 2);
  // c) status PJ
  const P0 = W0 + wN + 3, PJS = "'" + SHEET.PJ + "'!$I$5:$I";
  sh.getRange(P0, 14, 4, 2).setValues([['Status PJ', 'Jumlah'], ['Diterima', '=COUNTIF(' + PJS + ',"Diterima")'],
    ['Dalam pengiriman', '=COUNTIF(' + PJS + ',"Dalam pengiriman")'], ['Tidak kirim', '=COUNTIF(' + PJS + ',"Tidak kirim")']]);
  sh.getRange(P0, 14, 1, 2).setFontWeight('bold');
  sh.getRange(P0 + 4, 14, 1, 2).setValues([['Rata-rata lama kirim (hari)', "=IFERROR(AVERAGE('" + SHEET.PJ + "'!$H$5:$H),0)"]]);
  sh.getRange(P0 + 4, 15).setNumberFormat('0.0');
  const pjR = sh.getRange(P0, 14, 4, 2);
  // d) kualitas data
  const Q0 = P0 + 7;
  sh.getRange(Q0, 14, 5, 2).setValues([['Status data', 'Jumlah'], ['OK', '=COUNTIF(' + J_('S') + ',"OK")'], ['Info', '=COUNTIF(' + J_('S') + ',"Info")'],
    ['Dikoreksi', '=COUNTIF(' + J_('S') + ',"Dikoreksi")'], ['Perlu cek', '=COUNTIF(' + J_('S') + ',"Perlu cek")']]);
  sh.getRange(Q0, 14, 1, 2).setFontWeight('bold');
  const qR = sh.getRange(Q0, 14, 5, 2);
  sh.setColumnWidth(14, 260); sh.setColumnWidth(15, 130);

  // ---- grafik
  const bl = ss.getSheetByName(SHEET.BULANAN), nM = bl.getLastRow() - 5;
  const cTM = 3 + AKUN.filter((a) => a[3] === 'Masuk').length, cTK = cTM + 1 + out.length, cSA = cTK + 1;
  const pr = ss.getSheetByName(SHEET.PERIODE), nP = pr.getLastRow() - 5;
  const pq = ss.getSheetByName(SHEET.PENGAJUAN);

  const W = 960, H = 320;
  // 1. Arus kas bulanan (combo)
  sh.insertChart(sh.newChart().asComboChart()
    .addRange(bl.getRange(4, 1, nM + 1, 1)).addRange(bl.getRange(4, cTM, nM + 1, 1)).addRange(bl.getRange(4, cTK, nM + 1, 1)).addRange(bl.getRange(4, cSA, nM + 1, 1))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS).setNumHeaders(1)
    .setOption('title', 'Arus kas bulanan: masuk, keluar, dan saldo akhir')
    .setOption('series', { 0: { type: 'bars', color: '#2F6FB0' }, 1: { type: 'bars', color: '#E07B39' }, 2: { type: 'line', color: '#2E8B57', lineWidth: 3 } })
    .setOption('legend', { position: 'top' }).setOption('vAxis', { format: 'short' }).setOption('hAxis', { format: 'MMM yy' })
    .setOption('width', W).setOption('height', H).setPosition(7, 1, 0, 0).build());
  // 2. Komposisi pengeluaran
  sh.insertChart(sh.newChart().asBarChart().addRange(komp).setNumHeaders(1)
    .setOption('title', 'Ke mana uang keluar (total per akun)').setOption('legend', { position: 'none' }).setOption('colors', ['#E07B39'])
    .setOption('hAxis', { format: 'short' }).setOption('width', 470).setOption('height', H).setPosition(24, 1, 0, 0).build());
  // 3. Wilayah (tahun terpilih)
  sh.insertChart(sh.newChart().asBarChart().addRange(wilR).setNumHeaders(1)
    .setOption('title', 'Biaya ops bersih per wilayah (ganti tahun di sel kuning, kolom O)').setOption('legend', { position: 'none' }).setOption('colors', ['#2F6FB0'])
    .setOption('hAxis', { format: 'short' }).setOption('width', 480).setOption('height', 420).setPosition(24, 5, 10, 0).build());
  // 4. Transfer per periode
  sh.insertChart(sh.newChart().asLineChart().addRange(pr.getRange(4, 2, nP + 1, 1)).addRange(pr.getRange(4, 3, nP + 1, 1)).addRange(pr.getRange(4, 6, nP + 1, 1))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS).setNumHeaders(1)
    .setOption('title', 'Transfer operasional per periode dua mingguan').setOption('legend', { position: 'top' })
    .setOption('series', { 0: { targetAxisIndex: 0, color: '#2F6FB0', lineWidth: 2 }, 1: { targetAxisIndex: 1, color: '#9AA5B1', lineDashStyle: [4, 4] } })
    .setOption('vAxes', { 0: { format: 'short', title: 'Rupiah' }, 1: { title: 'Jumlah penerima' } })
    .setOption('width', W).setOption('height', H).setPosition(46, 1, 0, 0).build());
  // 5. Pengajuan vs ACC per agroman (pivot via QUERY)
  if (pq && pq.getLastRow() > 4) {
    const A0 = Q0 + 8;
    sh.getRange(A0, 14).setFormula("=QUERY('" + SHEET.PENGAJUAN + "'!A5:F,\"select B, sum(E), sum(F) where E > 0 and F is not null group by B label B 'Agroman', sum(E) 'Pengajuan', sum(F) 'ACC'\",0)");
    SpreadsheetApp.flush();
    const nA = Math.max(1, sh.getRange(A0, 14, 60, 1).getValues().filter((r) => r[0] !== '').length);
    sh.getRange(A0, 15, nA, 2).setNumberFormat(CFG.FORMAT_ANGKA);
    sh.insertChart(sh.newChart().asBarChart().addRange(sh.getRange(A0, 14, nA, 3)).setNumHeaders(1)
      .setOption('title', 'Pengajuan vs ACC per Agroman').setOption('legend', { position: 'top' }).setOption('colors', ['#B8C4D3', '#2F6FB0'])
      .setOption('hAxis', { format: 'short' }).setOption('width', 470).setOption('height', 420).setPosition(63, 1, 0, 0).build());
  }
  // 6. Status PJ & 7. kualitas data
  sh.insertChart(sh.newChart().asPieChart().addRange(pjR).setNumHeaders(1).setOption('title', 'Status berkas PJ')
    .setOption('colors', ['#2E8B57', '#E0A400', '#C0392B']).setOption('pieHole', 0.45)
    .setOption('width', 480).setOption('height', 210).setPosition(63, 5, 10, 0).build());
  sh.insertChart(sh.newChart().asPieChart().addRange(qR).setNumHeaders(1).setOption('title', 'Kualitas data Jurnal Kas')
    .setOption('colors', ['#2E8B57', '#7FA7D6', '#E0A400', '#C0392B']).setOption('pieHole', 0.45)
    .setOption('width', 480).setOption('height', 210).setPosition(74, 5, 10, 0).build());

  sh.getRange('A1:O200').setFontFamily(CFG.FONT);
  sh.setFrozenRows(5);
  sh.activate();
}

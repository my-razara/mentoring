/**
 * Dashboard Mentoring Ashabul Istiqomah
 * Backend Google Apps Script (database: Google Spreadsheet)
 *
 * Tampilan ada di GitHub (repo mentoring), file ini cuma API ke spreadsheet.
 * 1. Paste file ini ke Code.gs di Apps Script spreadsheet mentoring
 * 2. Hapus file HTML "Index" dari project Apps Script (udah nggak dipakai)
 * 3. Jalankan setupSpreadsheet() sekali kalau sheet belum lengkap
 * 4. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy URL /exec-nya ke config.js di repo GitHub
 */

// Kode buat masuk mode pengurus (yang boleh nambah/edit/hapus). Ganti sesuai selera.
const KODE_PENGURUS = 'istiqomah123';

const SHEETS = {
  program: 'Program',
  rundown: 'Rundown',
  persiapan: 'Persiapan',
  anggota: 'Anggota',
  profil: 'Profil',
  catatan: 'Catatan',
  dokumentasi: 'Dokumentasi',
  keuangan: 'Keuangan'
};

const HEADERS = {
  Program: ['ID', 'Nama Program', 'Kategori', 'Tanggal Mulai', 'Tanggal Selesai', 'Waktu', 'Lokasi', 'Penanggung Jawab', 'Status', 'Tujuan', 'Catatan', 'Diupdate'],
  Rundown: ['ID', 'ID Program', 'Mulai', 'Selesai', 'Kegiatan', 'PIC', 'Keterangan'],
  Persiapan: ['ID', 'ID Program', 'Item', 'PIC', 'Deadline', 'Selesai'],
  Anggota: ['Nama', 'Peran', 'Kontak', 'Aktif', 'Urutan', 'Foto'],
  Profil: ['ID', 'Bagian', 'Judul', 'Isi', 'Urutan'],
  Catatan: ['ID', 'Tanggal', 'Judul', 'Isi', 'Pencatat'],
  Dokumentasi: ['ID', 'ID Program', 'Judul', 'Link', 'Tanggal'],
  Keuangan: ['ID', 'Tanggal', 'Kas', 'Jenis', 'Keterangan', 'Jumlah']
};

const BAGIAN_LIST = ['Identitas', 'Pengertian', 'Tujuan', 'Program Kerja / Kegiatan Sosial'];

const KAS_LIST = ['Kas Mentoring', 'Kas Yatim'];
const JENIS_LIST = ['Masuk', 'Keluar'];

const STATUS_LIST = ['Rencana', 'Persiapan', 'Berjalan', 'Selesai', 'Batal'];
const KATEGORI_LIST = ['Mentoring Rutin', 'Olahraga', 'Nongkrong', 'Milad', 'Sosial', 'Keagamaan', 'Lainnya'];

/* ============================ API ============================ */
/*
 * Tampilan dashboard sekarang di GitHub Pages. Apps Script cuma jadi API:
 * - GET  ?action=getData           -> ambil semua data
 * - POST {fn, args} (text/plain)   -> jalanin satu fungsi dari daftar API_FUNGSI
 * Deploy sebagai Web app: Execute as Me, Who has access: Anyone.
 */

const API_FUNGSI = [
  'getData', 'cekKode',
  'simpanProgram', 'ubahStatus', 'hapusProgram',
  'simpanRundown', 'hapusRundown',
  'simpanPersiapan', 'togglePersiapan', 'hapusPersiapan',
  'simpanAnggota', 'hapusAnggota',
  'simpanProfil', 'hapusProfil',
  'simpanCatatan', 'hapusCatatan',
  'simpanKeuangan', 'hapusKeuangan',
  'simpanDokumentasi', 'hapusDokumentasi'
];

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _jalankan(fn, args) {
  if (API_FUNGSI.indexOf(fn) === -1) throw new Error('Fungsi ' + fn + ' tidak tersedia.');
  const f = globalThis[fn];
  if (typeof f !== 'function') throw new Error('Fungsi ' + fn + ' belum ada di Code.gs.');
  return f.apply(null, Array.isArray(args) ? args : []);
}

function doGet(e) {
  try {
    const fn = (e && e.parameter && e.parameter.action) || 'getData';
    if (fn !== 'getData') throw new Error('GET cuma buat getData.');
    return _json({ ok: true, data: getData() });
  } catch (err) {
    return _json({ ok: false, error: err.message || String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return _json({ ok: true, data: _jalankan(body.fn, body.args) });
  } catch (err) {
    return _json({ ok: false, error: err.message || String(err) });
  }
}

/* ============================ SETUP ============================ */

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _migrasiProfil(ss);

  Object.keys(HEADERS).forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]])
      .setFontWeight('bold').setBackground('#1E6F5C').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  });

  // format kolom waktu & tanggal sebagai teks biar ga diubah-ubah Sheets
  const prog = ss.getSheetByName(SHEETS.program);
  prog.getRange('D2:F').setNumberFormat('@');
  ss.getSheetByName(SHEETS.rundown).getRange('C2:D').setNumberFormat('@');
  ss.getSheetByName(SHEETS.persiapan).getRange('E2:E').setNumberFormat('@');

  // dropdown status & kategori
  prog.getRange('C2:C').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(KATEGORI_LIST, true).build());
  prog.getRange('I2:I').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS_LIST, true).build());

  // contoh data kalau masih kosong
  if (prog.getLastRow() < 2) {
    const p1 = 'PRG-' + Date.now().toString(36);
    const p2 = 'PRG-' + (Date.now() + 1).toString(36);
    prog.getRange(2, 1, 2, HEADERS.Program.length).setValues([
      [p1, 'Mentoring Pekanan', 'Mentoring Rutin', '2026-09-20', '', '16.00 - 17.30', 'Masjid Al-Istiqomah',
       'Razan', 'Berjalan', 'Menjaga rutinitas halaqah pekanan dan mempererat ukhuwah antaranggota.', '', new Date()],
      [p2, 'Kajian Akbar + Makan Bersama', 'Nongkrong', '2026-10-11', '', '08.00 - 11.30', 'Aula RW 04',
       'Razan', 'Rencana', 'Menghadirkan pemateri luar dan membuka kajian untuk teman di luar mentoring.', '', new Date()]
    ]);

    const rd = ss.getSheetByName(SHEETS.rundown);
    rd.getRange(2, 1, 4, HEADERS.Rundown.length).setValues([
      ['RD-1', p1, '16.00', '16.10', 'Pembukaan dan tilawah', 'Fajar', ''],
      ['RD-2', p1, '16.10', '16.40', 'Setoran hafalan', 'Razan', 'Tiap anggota satu halaman'],
      ['RD-3', p1, '16.40', '17.15', 'Materi pekanan', 'Pemateri', ''],
      ['RD-4', p1, '17.15', '17.30', 'Sharing dan doa penutup', 'Fajar', '']
    ]);

    const ps = ss.getSheetByName(SHEETS.persiapan);
    ps.getRange(2, 1, 3, HEADERS.Persiapan.length).setValues([
      ['CK-1', p2, 'Konfirmasi pemateri', 'Razan', '2026-09-25', 'Belum'],
      ['CK-2', p2, 'Pinjam aula dan sound system', 'Fajar', '2026-09-28', 'Belum'],
      ['CK-3', p2, 'Buat poster dan sebar undangan', 'Razan', '2026-10-01', 'Belum']
    ]);

    ss.getSheetByName(SHEETS.anggota).getRange(2, 1, 2, HEADERS.Anggota.length).setValues([
      ['Razan', 'Ketua', '', 'Ya', 1, ''],
      ['Fajar', 'Anggota', '', 'Ya', 2, '']
    ]);
  }

  const pf = ss.getSheetByName(SHEETS.profil);
  if (pf.getLastRow() < 2) {
    pf.getRange(2, 1, 6, 5).setValues([
      [_newId('PF'), 'Identitas', 'Logo', '', 0],
      [_newId('PF'), 'Identitas', 'Nama Kelompok', 'Mentoring Ashabul Istiqomah', 1],
      [_newId('PF'), 'Identitas', 'Periode', '2026 / 2027', 2],
      [_newId('PF'), 'Identitas', 'Deskripsi', 'Kelompok mentoring yang fokus menjaga rutinitas ibadah dan mempererat ukhuwah antaranggota.', 3],
      [_newId('PF'), 'Identitas', 'Jadwal Rutin', 'Setiap Sabtu, 16.00 sampai 17.30', 4],
      [_newId('PF'), 'Identitas', 'Tempat', 'Masjid Al-Istiqomah', 5]
    ]);
    pf.setColumnWidth(4, 420);
  }

  const ct = ss.getSheetByName(SHEETS.catatan);
  ct.getRange('B2:B').setNumberFormat('@');
  ct.setColumnWidth(4, 420);

  const kg = ss.getSheetByName(SHEETS.keuangan);
  kg.getRange('B2:B').setNumberFormat('@');
  kg.getRange('C2:C').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(KAS_LIST, true).build());
  kg.getRange('D2:D').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(JENIS_LIST, true).build());
  kg.getRange('F2:F').setNumberFormat('#,##0');

  return 'Setup selesai. Lanjut deploy web app.';
}

/**
 * Ngerapihin struktur sheet ke bentuk sekarang:
 * - Profil: ID, Bagian, Judul, Isi, Urutan
 * - Catatan: ID, Tanggal, Judul, Isi, Pencatat
 * Sheet Konten lama digabung ke Profil, dan baris Catatan Pertemuan yang
 * sempat nyampur di Profil dikembalikan ke sheet Catatan. Aman diulang.
 */
function _migrasiProfil(ss) {
  const pf = ss.getSheetByName(SHEETS.profil);
  if (!pf) return;

  const kt = ss.getSheetByName('Konten');
  const kolom = Math.max(pf.getLastColumn(), 1);
  const head = pf.getRange(1, 1, 1, kolom).getValues()[0].map(function (h) { return String(h).trim(); });
  const bentukLama = head[0] === 'Kunci';
  const bentukGabung = head[5] === 'Tanggal' && head[6] === 'Pencatat';
  if (!bentukLama && !bentukGabung && !kt) return;

  const barisProfil = [];
  const barisCatatan = [];

  if (pf.getLastRow() > 1) {
    const lebar = bentukLama ? 3 : (bentukGabung ? 7 : 5);
    pf.getRange(2, 1, pf.getLastRow() - 1, lebar).getValues().forEach(function (r) {
      if (bentukLama) {
        if (String(r[0]).trim()) barisProfil.push(['', 'Identitas', String(r[0]).trim(), String(r[1] || '')]);
        return;
      }
      if (!String(r[2]).trim()) return;
      const bagian = String(r[1] || 'Lainnya');
      if (bagian.toLowerCase().replace(/[^a-z]/g, '') === 'catatanpertemuan') {
        barisCatatan.push([String(r[0] || ''), _cell(r[5]), String(r[2]), String(r[3] || ''), String(r[6] || '')]);
      } else {
        barisProfil.push([String(r[0] || ''), bagian, String(r[2]), String(r[3] || '')]);
      }
    });
  }

  if (kt && kt.getLastRow() > 1) {
    kt.getRange(2, 1, kt.getLastRow() - 1, 5).getValues().forEach(function (r) {
      if (String(r[2]).trim()) barisProfil.push([String(r[0] || ''), String(r[1] || 'Lainnya'), String(r[2]), String(r[3] || '')]);
    });
  }

  pf.clear();
  pf.getRange(1, 1, 1, 5).setValues([HEADERS.Profil]);
  if (barisProfil.length) {
    pf.getRange(2, 1, barisProfil.length, 5).setValues(barisProfil.map(function (r, i) {
      return [r[0] || _newId('PF'), r[1], r[2], r[3], i + 1];
    }));
  }

  if (barisCatatan.length) {
    let ct = ss.getSheetByName(SHEETS.catatan);
    if (!ct) ct = ss.insertSheet(SHEETS.catatan);
    if (ct.getLastRow() < 1) ct.getRange(1, 1, 1, 5).setValues([HEADERS.Catatan]);
    const adaId = ct.getLastRow() > 1
      ? ct.getRange(2, 1, ct.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); })
      : [];
    const baru = barisCatatan.filter(function (r) { return adaId.indexOf(r[0]) === -1; });
    if (baru.length) {
      ct.getRange(ct.getLastRow() + 1, 1, baru.length, 5).setValues(baru.map(function (r) {
        return [r[0] || _newId('CT'), r[1], r[2], r[3], r[4]];
      }));
      ct.getRange(2, 2, ct.getLastRow() - 1, 1).setNumberFormat('@');
    }
  }

  if (kt) ss.deleteSheet(kt);
}

/**
 * Ngisi Profil, Konten, dan Program kerja dari hasil rapat evaluasi mentoring.
 * Aman dijalankan berkali-kali: baris yang judulnya udah ada bakal dilewati.
 * Jalanin sekali dari editor Apps Script, sama kayak setupSpreadsheet().
 */
function isiDataRapat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSpreadsheet();
  const catatan = [];

  /* ---- Profil: identitas, pengertian, metode, organisasi ---- */
  const pf = ss.getSheetByName(SHEETS.profil);
  const isiProfil = [
    ['Identitas', 'Motto',
      'Berproses, bertumbuh dan berkembang bersama keluarga mentoring Ashabul Istiqomah. Dari ane untuk mentoring yang hebat dan berdampak.'],
    ['Pengertian', 'Apa itu mentoring',
      'Mentoring adalah proses bimbingan, pendidikan, pengetahuan, dan penanaman nilai karakter dari seorang mentor melalui dinamika bersama anggota, untuk mendukung pengembangan pribadi, sosial, belajar karier, dan agama.'],
    ['Pengertian', 'Bentuk layanan',
      'Ada dua bentuk layanan. Pertama liqo atau pertemuan kelompok bersama. Kedua konseling dan konsultasi individu, menyesuaikan kebutuhan masing-masing anggota.'],
    ['Metode', 'Jadwal pertemuan',
      'Mulai Oktober 2026 pertemuan diadakan di momen tertentu: libur, tanggal merah, masa semesteran, atau akhir tahun.\n\nRiwayat sebelumnya: sepekan sekali dan dua pekan sekali saat masa sekolah di SMKN 22, lalu sebulan sekali dan dua sampai tiga bulan sekali setelah lulus.\n\nFrekuensi finalnya masih perlu diskusi lanjut.'],
    ['Organisasi', 'Struktur pengurus',
      'Divisi-divisi mentoring dibubarkan dan diganti pengurus inti yang terdiri dari Ketua, Sekretaris, dan Bendahara.'],
    ['Organisasi', 'Kas mentoring',
      'Ada dua kas: Kas Mentoring dan Kas Yatim. Rekap saldo per September 2026 dipegang Husein dan Latif.'],
    ['Organisasi', 'Kendala yang dibahas',
      'Delapan kendala hasil evaluasi: kurang kompak dan solid, sulit mengatur waktu, sulit mengatur tempat, sikap mengabaikan dan menyepelekan, belum punya tujuan yang jelas, drama hubungan, drama perkuliahan, dan drama pekerjaan.\n\nSolusi tiap poin masih perlu diskusi lanjut.'],
    ['Organisasi', 'Tindak lanjut',
      '1. Berdoa bersama meminta kekuatan istiqomah dan petunjuk dari Allah SWT dan Rasulullah SAW.\n2. Berjanji dan berkomitmen bersama anggota mentoring.\n3. Saling meminta maaf dan menjaga komunikasi untuk berproses, bertumbuh, dan berkembang bersama di Mentoring Ashabul Istiqomah.']
  ];
  const judulAda = pf.getLastRow() > 1
    ? pf.getRange(2, 3, pf.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim().toLowerCase(); })
    : [];
  isiProfil.forEach(function (r) {
    if (judulAda.indexOf(r[1].toLowerCase()) === -1) {
      const baris = pf.getLastRow() + 1;
      pf.getRange(baris, 1, 1, 5).setValues([[_newId('PF'), r[0], r[1], r[2], baris - 1]]);
      catatan.push('Profil: ' + r[1]);
    }
  });

  /* ---- Program kerja ---- */
  const prog = ss.getSheetByName(SHEETS.program);
  const ps = ss.getSheetByName(SHEETS.persiapan);
  const isiProgram = [
    {
      nama: 'Santunan dan Open Donasi Yatim Dhuafa',
      kategori: 'Sosial',
      tujuan: 'Menyalurkan santunan dan donasi dari anggota mentoring untuk yatim dan dhuafa.',
      catatan: 'Waktu pelaksanaan: momen Lebaran yatim dan bulan Ramadhan. Teknis, target dana, dan penerima masih perlu diskusi lanjut.',
      persiapan: ['Tentukan target dana dan calon penerima', 'Buka open donasi ke anggota', 'Rekap donasi masuk', 'Salurkan santunan dan dokumentasikan']
    },
    {
      nama: 'Kegiatan Ramadhan',
      kategori: 'Keagamaan',
      tujuan: 'Mengisi bulan Ramadhan dengan rangkaian kegiatan bersama anggota mentoring.',
      catatan: 'Waktu: bulan Ramadhan. Rincian tiap kegiatan masih perlu diskusi lanjut.',
      persiapan: ['Buka bersama', 'ODOZ, one day one juz', 'Bagi takjil', 'Sedekah', 'Silaturahmi Idul Fitri']
    },
    {
      nama: 'PHBI, Peringatan Hari Besar Islam',
      kategori: 'Keagamaan',
      tujuan: 'Memperingati hari besar Islam bersama anggota sepanjang tahun.',
      catatan: 'Waktu menyesuaikan kalender Hijriah. Teknis tiap agenda masih perlu diskusi lanjut.',
      persiapan: ['Buka bersama puasa sunnah Idul Adha dan bakar-bakar', 'Nuzulul Quran atau Lailatul Qadar', 'Isra Miraj', 'Maulid Nabi', 'Muharram', 'Rouha dan ziarah menjelang Ramadhan']
    }
  ];
  const namaAda = prog.getLastRow() > 1
    ? prog.getRange(2, 2, prog.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim().toLowerCase(); })
    : [];

  isiProgram.forEach(function (p) {
    if (namaAda.indexOf(p.nama.toLowerCase()) > -1) return;
    const id = _newId('PRG');
    prog.getRange(prog.getLastRow() + 1, 1, 1, HEADERS.Program.length).setValues([[
      id, p.nama, p.kategori, '', '', '', '', '', 'Rencana', p.tujuan, p.catatan,
      Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd HH:mm')
    ]]);
    p.persiapan.forEach(function (item) {
      ps.getRange(ps.getLastRow() + 1, 1, 1, HEADERS.Persiapan.length)
        .setValues([[_newId('CK'), id, item, '', '', 'Belum']]);
    });
    catatan.push('Program: ' + p.nama);
  });

  /* ---- Tambahan buat Milad yang udah ada ---- */
  const idMilad = _cariMilad(prog);
  if (idMilad) {
    const sudah = ps.getLastRow() > 1
      ? ps.getRange(2, 3, ps.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim().toLowerCase(); })
      : [];
    ['Tentukan tempat final: vila, camping, atau piknik ke Telaga Saat',
     'Susun anggaran biaya',
     'Siapkan hadiah untuk tukar kado',
     'Konfirmasi kehadiran anggota',
     'Siapkan perlengkapan acara'].forEach(function (item) {
      if (sudah.indexOf(item.toLowerCase()) === -1) {
        ps.getRange(ps.getLastRow() + 1, 1, 1, HEADERS.Persiapan.length)
          .setValues([[_newId('CK'), idMilad, item, '', '', 'Belum']]);
        catatan.push('Persiapan milad: ' + item);
      }
    });

    const rd = ss.getSheetByName(SHEETS.rundown);
    const adaRundown = rd.getLastRow() > 1 &&
      rd.getRange(2, 2, rd.getLastRow() - 1, 1).getValues()
        .some(function (r) { return String(r[0]).trim() === idMilad; });
    if (!adaRundown) {
      ['Liqo dan sharing', 'Games', 'Makan bersama', 'Tukar hadiah'].forEach(function (keg) {
        rd.getRange(rd.getLastRow() + 1, 1, 1, HEADERS.Rundown.length)
          .setValues([[_newId('RD'), idMilad, '', '', keg, '', '']]);
      });
      catatan.push('Rundown milad: 4 acara, jamnya diisi sendiri');
    }
  }

  return catatan.length ? 'Ditambahkan:\n' + catatan.join('\n') : 'Semua data sudah ada, nggak ada yang ditambah.';
}

function _cariMilad(prog) {
  if (prog.getLastRow() < 2) return '';
  const rows = prog.getRange(2, 1, prog.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase().indexOf('milad') > -1) return String(rows[i][0]);
  }
  return '';
}

/* ============================ HELPERS ============================ */
function _sheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" belum ada. Jalankan setupSpreadsheet() dulu.');
  return sheet;
}

function _tz() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Jakarta';
}

function _cell(value) {
  if (value instanceof Date) return Utilities.formatDate(value, _tz(), 'yyyy-MM-dd');
  return value === null || value === undefined ? '' : String(value).trim();
}

function _read(name) {
  const sheet = _sheet(name);
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const head = values.shift().map(function (h) { return String(h).trim(); });
  return values
    .map(function (row, i) {
      const obj = { _row: i + 2 };
      head.forEach(function (h, c) { if (h) obj[h] = _cell(row[c]); });
      return obj;
    })
    .filter(function (obj) { return head.some(function (h) { return h && obj[h] !== ''; }); });
}

function _findRow(name, id) {
  const sheet = _sheet(name);
  const ids = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) return i + 1;
  }
  return -1;
}

function _newId(prefix) {
  return prefix + '-' + Date.now().toString(36) + Math.floor(Math.random() * 90 + 10);
}

function _guard(kode) {
  if (String(kode || '').trim() !== KODE_PENGURUS) {
    throw new Error('Kode pengurus salah. Data tidak diubah.');
  }
}

/* ============================ READ ============================ */

function getData() {
  _autoSelesai();
  return {
    program: _read(SHEETS.program).sort(function (a, b) {
      return String(a['Tanggal Mulai'] || '9999').localeCompare(String(b['Tanggal Mulai'] || '9999'));
    }),
    rundown: _read(SHEETS.rundown).sort(function (a, b) {
      return String(a.Mulai || '').localeCompare(String(b.Mulai || ''));
    }),
    persiapan: _read(SHEETS.persiapan),
    anggota: _read(SHEETS.anggota).sort(function (a, b) {
      const ua = Number(a.Urutan) || 999, ub = Number(b.Urutan) || 999;
      return ua === ub ? String(a.Nama).localeCompare(String(b.Nama)) : ua - ub;
    }),
    profil: _read(SHEETS.profil).sort(function (a, b) {
      return (Number(a.Urutan) || 999) - (Number(b.Urutan) || 999);
    }),
    catatan: _read(SHEETS.catatan).sort(function (a, b) {
      return String(b.Tanggal || '').localeCompare(String(a.Tanggal || ''));
    }),
    dokumentasi: _read(SHEETS.dokumentasi).sort(function (a, b) {
      return String(b.Tanggal || '').localeCompare(String(a.Tanggal || ''));
    }),
    keuangan: _read(SHEETS.keuangan).sort(function (a, b) {
      return String(b.Tanggal || '').localeCompare(String(a.Tanggal || ''));
    }),
    opsi: { status: STATUS_LIST, kategori: KATEGORI_LIST, kas: KAS_LIST, jenis: JENIS_LIST, bagian: BAGIAN_LIST }
  };
}

/**
 * Program yang tanggal terakhirnya udah lewat (minimal kemarin) otomatis jadi Selesai.
 * Status Batal dan Selesai nggak disentuh.
 */
function _autoSelesai() {
  const lock = LockService.getScriptLock();
  const sudahPegang = lock.hasLock();
  if (!sudahPegang && !lock.tryLock(5000)) return;
  try {
    const sheet = _sheet(SHEETS.program);
    const n = sheet.getLastRow() - 1;
    if (n < 1) return;

    const hariIni = Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd');
    const stempel = Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd HH:mm');
    const nilai = sheet.getRange(2, 1, n, HEADERS.Program.length).getValues();
    const status = nilai.map(function (r) { return [r[8]]; });
    const update = nilai.map(function (r) { return [r[11]]; });
    let berubah = false;

    nilai.forEach(function (r, i) {
      const s = String(r[8]).trim();
      if (s === 'Selesai' || s === 'Batal') return;
      const akhir = _cell(r[4]) || _cell(r[3]);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(akhir)) return;
      if (akhir < hariIni) {
        status[i][0] = 'Selesai';
        update[i][0] = stempel;
        berubah = true;
      }
    });

    if (berubah) {
      sheet.getRange(2, 9, n, 1).setValues(status);
      sheet.getRange(2, 12, n, 1).setValues(update);
    }
  } finally {
    if (!sudahPegang) lock.releaseLock();
  }
}

function cekKode(kode) {
  return String(kode || '').trim() === KODE_PENGURUS;
}

/* ============================ WRITE: PROGRAM ============================ */

function simpanProgram(payload, kode) {
  _guard(kode);
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = _sheet(SHEETS.program);
    const baris = [
      data.id || _newId('PRG'),
      data.nama || '',
      data.kategori || '',
      data.tanggalMulai || '',
      data.tanggalSelesai || '',
      data.waktu || '',
      data.lokasi || '',
      data.pj || '',
      data.status || 'Rencana',
      data.tujuan || '',
      data.catatan || '',
      Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd HH:mm')
    ];
    let row = data.id ? _findRow(SHEETS.program, data.id) : -1;
    if (row < 0) row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, baris.length).setValues([baris]);
    sheet.getRange(row, 4, 1, 3).setNumberFormat('@');
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function ubahStatus(id, status, kode) {
  _guard(kode);
  const row = _findRow(SHEETS.program, id);
  if (row < 0) throw new Error('Program tidak ditemukan.');
  const sheet = _sheet(SHEETS.program);
  sheet.getRange(row, 9).setValue(status);
  sheet.getRange(row, 12).setValue(Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd HH:mm'));
  return getData();
}

function hapusProgram(id, kode) {
  _guard(kode);
  const row = _findRow(SHEETS.program, id);
  if (row > 0) _sheet(SHEETS.program).deleteRow(row);
  _hapusAnak(SHEETS.rundown, id);
  _hapusAnak(SHEETS.persiapan, id);
  return getData();
}

function _hapusAnak(sheetName, idProgram) {
  const sheet = _sheet(sheetName);
  if (sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() === String(idProgram).trim()) sheet.deleteRow(i + 2);
  }
}

/* ============================ WRITE: RUNDOWN ============================ */

function simpanRundown(payload, kode) {
  _guard(kode);
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  const sheet = _sheet(SHEETS.rundown);
  const baris = [
    data.id || _newId('RD'),
    data.idProgram,
    data.mulai || '',
    data.selesai || '',
    data.kegiatan || '',
    data.pic || '',
    data.keterangan || ''
  ];
  let row = data.id ? _findRow(SHEETS.rundown, data.id) : -1;
  if (row < 0) row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, baris.length).setValues([baris]);
  sheet.getRange(row, 3, 1, 2).setNumberFormat('@');
  return getData();
}

function hapusRundown(id, kode) {
  _guard(kode);
  const row = _findRow(SHEETS.rundown, id);
  if (row > 0) _sheet(SHEETS.rundown).deleteRow(row);
  return getData();
}

/* ============================ WRITE: PERSIAPAN ============================ */

// Daftar persiapan terbuka buat semua anggota: nambah, edit, centang, dan hapus
// nggak butuh kode pengurus. Program, rundown, dan anggota tetap dikunci.
function simpanPersiapan(payload) {
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  const sheet = _sheet(SHEETS.persiapan);
  const baris = [
    data.id || _newId('CK'),
    data.idProgram,
    data.item || '',
    data.pic || '',
    data.deadline || '',
    data.selesai ? 'Selesai' : 'Belum'
  ];
  let row = data.id ? _findRow(SHEETS.persiapan, data.id) : -1;
  if (row < 0) row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, baris.length).setValues([baris]);
  sheet.getRange(row, 5).setNumberFormat('@');
  return getData();
}

// Centang persiapan sengaja tidak butuh kode pengurus, biar anggota ikut update.
// Nambah, edit, dan hapus item tetap cuma bisa lewat mode pengurus.
function togglePersiapan(id, selesai) {
  const row = _findRow(SHEETS.persiapan, id);
  if (row > 0) _sheet(SHEETS.persiapan).getRange(row, 6).setValue(selesai ? 'Selesai' : 'Belum');
  return getData();
}

function hapusPersiapan(id) {
  const row = _findRow(SHEETS.persiapan, id);
  if (row > 0) _sheet(SHEETS.persiapan).deleteRow(row);
  return getData();
}

/* ============================ WRITE: ANGGOTA ============================ */

function simpanAnggota(payload, kode) {
  _guard(kode);
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  const nama = String(data.nama || '').trim();
  if (!nama) throw new Error('Nama anggota belum diisi.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = _sheet(SHEETS.anggota);
    const baris = [
      nama,
      data.peran || 'Anggota',
      data.kontak || '',
      data.aktif === false ? 'Tidak' : 'Ya',
      data.urutan === '' || data.urutan === undefined ? '' : Number(data.urutan),
      data.foto || ''
    ];
    let row = Number(data.row || 0);
    if (!row || row < 2 || row > sheet.getLastRow()) row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, baris.length).setValues([baris]);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function hapusAnggota(row, kode) {
  _guard(kode);
  const sheet = _sheet(SHEETS.anggota);
  const r = Number(row);
  if (r >= 2 && r <= sheet.getLastRow()) sheet.deleteRow(r);
  return getData();
}

/* ============================ WRITE: PROFIL ============================ */

function simpanProfil(payload, kode) {
  _guard(kode);
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  if (!String(data.judul || '').trim()) throw new Error('Judul belum diisi.');
  const sheet = _sheet(SHEETS.profil);
  const baris = [
    data.id || _newId('PF'),
    data.bagian || 'Lainnya',
    String(data.judul).trim(),
    String(data.isi || ''),
    data.urutan === '' || data.urutan === undefined ? '' : Number(data.urutan)
  ];
  let row = data.id ? _findRow(SHEETS.profil, data.id) : -1;
  if (row < 0) row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, baris.length).setValues([baris]);
  return getData();
}

function hapusProfil(id, kode) {
  _guard(kode);
  const row = _findRow(SHEETS.profil, id);
  if (row > 0) _sheet(SHEETS.profil).deleteRow(row);
  return getData();
}

/* ============================ WRITE: CATATAN PERTEMUAN ============================ */

function simpanCatatan(payload, kode) {
  _guard(kode);
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  if (!String(data.judul || '').trim()) throw new Error('Judul pertemuan belum diisi.');
  const sheet = _sheet(SHEETS.catatan);
  const baris = [
    data.id || _newId('CT'),
    data.tanggal || Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd'),
    String(data.judul).trim(),
    String(data.isi || ''),
    String(data.pencatat || '').trim()
  ];
  let row = data.id ? _findRow(SHEETS.catatan, data.id) : -1;
  if (row < 0) row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, baris.length).setValues([baris]);
  sheet.getRange(row, 2).setNumberFormat('@');
  return getData();
}

function hapusCatatan(id, kode) {
  _guard(kode);
  const row = _findRow(SHEETS.catatan, id);
  if (row > 0) _sheet(SHEETS.catatan).deleteRow(row);
  return getData();
}

/* ============================ WRITE: KEUANGAN ============================ */

function simpanKeuangan(payload, kode) {
  _guard(kode);
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  if (!String(data.keterangan || '').trim()) throw new Error('Keterangan belum diisi.');
  const jumlah = Number(data.jumlah);
  if (!jumlah || jumlah <= 0) throw new Error('Jumlah harus lebih dari nol.');

  const sheet = _sheet(SHEETS.keuangan);
  const baris = [
    data.id || _newId('KU'),
    data.tanggal || Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd'),
    KAS_LIST.indexOf(data.kas) > -1 ? data.kas : KAS_LIST[0],
    JENIS_LIST.indexOf(data.jenis) > -1 ? data.jenis : JENIS_LIST[0],
    String(data.keterangan).trim(),
    jumlah
  ];
  let row = data.id ? _findRow(SHEETS.keuangan, data.id) : -1;
  if (row < 0) row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, baris.length).setValues([baris]);
  sheet.getRange(row, 2).setNumberFormat('@');
  return getData();
}

function hapusKeuangan(id, kode) {
  _guard(kode);
  const row = _findRow(SHEETS.keuangan, id);
  if (row > 0) _sheet(SHEETS.keuangan).deleteRow(row);
  return getData();
}

/* ============================ WRITE: DOKUMENTASI ============================ */

// Nambah dokumentasi terbuka buat semua anggota.
function simpanDokumentasi(payload) {
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  if (!String(data.judul || '').trim()) throw new Error('Judul dokumentasi belum diisi.');
  if (!/^https?:\/\//i.test(String(data.link || ''))) throw new Error('Link harus diawali http:// atau https://');
  const sheet = _sheet(SHEETS.dokumentasi);
  const baris = [
    data.id || _newId('DK'),
    data.idProgram || '',
    String(data.judul).trim(),
    String(data.link).trim(),
    data.tanggal || Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd')
  ];
  let row = data.id ? _findRow(SHEETS.dokumentasi, data.id) : -1;
  if (row < 0) row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, baris.length).setValues([baris]);
  sheet.getRange(row, 5).setNumberFormat('@');
  return getData();
}

// Hapus dokumentasi tetap dikunci supaya arsip nggak hilang nggak sengaja.
function hapusDokumentasi(id, kode) {
  _guard(kode);
  const row = _findRow(SHEETS.dokumentasi, id);
  if (row > 0) _sheet(SHEETS.dokumentasi).deleteRow(row);
  return getData();
}

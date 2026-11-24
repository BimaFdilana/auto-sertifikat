/**
 * AUTO SERTIFIKAT — Webinar DWP Kabupaten Bengkalis 2026
 *
 * Alur: Google Form -> Google Sheets -> (trigger) -> salin template Slides
 *       -> ganti {{NAMA}} -> simpan PDF di Drive -> (opsional) kirim email.
 *
 * Cara pakai singkat (detail lengkap ada di README.md):
 *   1. Isi TEMPLATE_ID di bawah (ID Google Slides template sertifikat).
 *   2. Jalankan fungsi setup() SATU KALI -> membuat Form + Spreadsheet + trigger.
 *   3. Lihat link Form & Sheet di menu View > Logs (atau Execution log).
 */

// ============================== KONFIGURASI ==============================
var CONFIG = {
  // ID Google Slides template (bagian panjang di URL: /presentation/d/<ID>/edit)
  TEMPLATE_ID: 'ISI_ID_SLIDES_TEMPLATE_DI_SINI',

  // Teks placeholder di Slides yang akan diganti dengan nama peserta
  PLACEHOLDER: '{{NAMA}}',

  // Judul form
  FORM_TITLE: 'DAFTAR HADIR DIRGAHAYU DWP KE-27 WEBINAR DHARMA WANITA PERSATUAN KABUPATEN BENGKALIS TAHUN 2026',

  // Nama folder Drive tempat PDF sertifikat disimpan (dibuat otomatis kalau belum ada)
  OUTPUT_FOLDER_NAME: 'Sertifikat DWP Bengkalis 2026',

  // Pesan di halaman konfirmasi Form (link halaman unduh ditambahkan otomatis oleh updateFormConfirmation)
  CONFIRMATION_TEXT: 'Terima kasih, kehadiran Anda sudah tercatat. Klik tautan di bawah untuk mengunduh sertifikat (masukkan email yang sama dengan yang Anda pakai di form).',

  // Kirim sertifikat ke email peserta secara otomatis?
  SEND_EMAIL: true,
  EMAIL_SUBJECT: 'Sertifikat Webinar Dharma Wanita Persatuan Kabupaten Bengkalis 2026',

  // Format nama: 'asis' (apa adanya), 'upper' (HURUF BESAR), 'title' (Huruf Awal Besar).
  // Disarankan 'asis' karena nama bergelar (S.Pd., M.Si.) bisa rusak oleh 'title'.
  NAME_CASE: 'asis',

  // Kalau nama lebih panjang dari ini, ukuran font diperkecil supaya tidak keluar garis
  LONG_NAME_LENGTH: 32,
  LONG_NAME_FONT_SIZE: 22
};

// Judul kolom di Sheet (harus sama dengan pertanyaan di Form)
var COL = {
  NAMA: 'NAMA LENGKAP',
  INSTANSI: 'INSTANSI',
  EMAIL: 'Email Address',
  STATUS: 'Status Sertifikat',
  LINK: 'Link Sertifikat (PDF)'
};

// ============================== SETUP (jalankan sekali) ==============================
function setup() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) {
    throw new Error('Setup sudah pernah dijalankan. Hapus Script Properties (Project Settings) kalau ingin mengulang.');
  }

  var form = FormApp.create(CONFIG.FORM_TITLE);
  form.setDescription('Silakan isi daftar hadir. Sertifikat akan dibuat otomatis sesuai nama yang Anda tulis, jadi pastikan penulisan nama sudah benar (termasuk gelar).');
  form.setCollectEmail(true);          // email dipakai untuk mengirim sertifikat
  form.setAllowResponseEdits(false);   // nama tidak bisa diubah setelah kirim (sertifikat sudah terbit)
  form.setLimitOneResponsePerUser(false);

  form.addTextItem().setTitle(COL.NAMA).setRequired(true);
  form.addTextItem().setTitle(COL.INSTANSI).setRequired(true);

  var ss = SpreadsheetApp.create('Daftar Hadir DWP Bengkalis 2026 (Respons)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();

  // Trigger otomatis setiap ada form yang masuk
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();

  props.setProperties({
    FORM_ID: form.getId(),
    SHEET_ID: ss.getId()
  });

  form.setConfirmationMessage(CONFIG.CONFIRMATION_TEXT);

  Logger.log('LANGKAH BERIKUTNYA: Deploy > New deployment > Web app, lalu jalankan updateFormConfirmation().');
  Logger.log('FORM (bagikan ke peserta): ' + form.getPublishedUrl());
  Logger.log('FORM (edit): ' + form.getEditUrl());
  Logger.log('SPREADSHEET: ' + ss.getUrl());
}

// ============================== TRIGGER ==============================
function onFormSubmit(e) {
  // Kunci supaya banyak peserta submit bersamaan tidak saling tabrakan
  var lock = LockService.getScriptLock();
  lock.waitLock(120000);
  try {
    var sheet = e.range.getSheet();
    processRow_(sheet, e.range.getRow());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Jalankan manual kalau ada baris yang gagal / belum diproses
 * (kolom Status Sertifikat kosong atau berisi "GAGAL").
 */
function generateMissingCertificates() {
  var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var sheet = SpreadsheetApp.openById(ssId).getSheets()[0];
  var lock = LockService.getScriptLock();
  lock.waitLock(120000);
  try {
    var headers = ensureColumns_(sheet);
    var last = sheet.getLastRow();
    for (var r = 2; r <= last; r++) {
      var status = String(sheet.getRange(r, headers[COL.STATUS]).getValue());
      if (status === '' || status.indexOf('GAGAL') === 0) {
        processRow_(sheet, r);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ============================== INTI PROSES ==============================
function processRow_(sheet, row) {
  var headers = ensureColumns_(sheet);
  var statusCell = sheet.getRange(row, headers[COL.STATUS]);
  var linkCell = sheet.getRange(row, headers[COL.LINK]);

  try {
    var rawName = sheet.getRange(row, headers[COL.NAMA]).getValue();
    var name = formatName_(rawName);
    if (!name) throw new Error('Nama kosong');

    var pdfFile = buildCertificate_(name);
    linkCell.setValue('https://drive.google.com/uc?export=download&id=' + pdfFile.getId());

    var email = headers[COL.EMAIL] ? String(sheet.getRange(row, headers[COL.EMAIL]).getValue()).trim() : '';
    if (CONFIG.SEND_EMAIL && email) {
      MailApp.sendEmail({
        to: email,
        subject: CONFIG.EMAIL_SUBJECT,
        htmlBody: 'Yth. <b>' + escapeHtml_(name) + '</b>,<br><br>' +
          'Terima kasih telah mengikuti Webinar Dharma Wanita Persatuan Kabupaten Bengkalis Tahun 2026. ' +
          'Sertifikat Anda terlampir pada email ini.<br><br>Salam,<br>Panitia Webinar DWP Kabupaten Bengkalis',
        attachments: [pdfFile.getAs(MimeType.PDF)]
      });
      statusCell.setValue('TERKIRIM ' + new Date().toLocaleString('id-ID'));
    } else {
      statusCell.setValue('SELESAI (PDF dibuat)');
    }
  } catch (err) {
    statusCell.setValue('GAGAL: ' + err.message);
    console.error('Baris ' + row + ': ' + err.stack);
  }
}

function buildCertificate_(name) {
  var folder = getOutputFolder_();
  var template = DriveApp.getFileById(CONFIG.TEMPLATE_ID);
  var copy = template.makeCopy('tmp - ' + name, folder);

  try {
    var pres = SlidesApp.openById(copy.getId());
    var replaced = false;

    pres.getSlides().forEach(function (slide) {
      slide.getPageElements().forEach(function (el) {
        if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
        var shape = el.asShape();
        if (shape.getText().asString().indexOf(CONFIG.PLACEHOLDER) === -1) return;

        shape.getText().replaceAllText(CONFIG.PLACEHOLDER, name);
        if (name.length > CONFIG.LONG_NAME_LENGTH) {
          shape.getText().getTextStyle().setFontSize(CONFIG.LONG_NAME_FONT_SIZE);
        }
        replaced = true;
      });
    });

    if (!replaced) {
      throw new Error('Placeholder ' + CONFIG.PLACEHOLDER + ' tidak ditemukan di template Slides');
    }
    pres.saveAndClose();

    var pdfBlob = DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF)
      .setName('Sertifikat - ' + name + '.pdf');
    var pdf = folder.createFile(pdfBlob);
    // Agar peserta bisa mengunduh lewat link tanpa login
    pdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return pdf;
  } finally {
    copy.setTrashed(true); // buang salinan Slides sementara
  }
}

// ============================== HALAMAN UNDUH (Web App) ==============================
function doGet() {
  return HtmlService.createHtmlOutput(DOWNLOAD_PAGE_)
    .setTitle('Unduh Sertifikat DWP Bengkalis 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Dipanggil dari halaman unduh. Hanya mengembalikan data milik email yang diketik. */
function findCertificate(email) {
  email = String(email || '').trim().toLowerCase();
  if (!/^\S+@\S+\.\S+$/.test(email)) return { state: 'invalid' };

  var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var sheet = SpreadsheetApp.openById(ssId).getSheets()[0];
  var last = sheet.getLastRow();
  var h = readHeaders_(sheet);
  if (last < 2 || !h[COL.EMAIL]) return { state: 'notfound' };

  var data = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (var i = data.length - 1; i >= 0; i--) { // baris terbaru dulu
    var row = data[i];
    if (String(row[h[COL.EMAIL] - 1]).trim().toLowerCase() !== email) continue;
    var link = h[COL.LINK] ? String(row[h[COL.LINK] - 1]) : '';
    var status = h[COL.STATUS] ? String(row[h[COL.STATUS] - 1]) : '';
    if (link) return { state: 'ready', url: link, name: String(row[h[COL.NAMA] - 1]) };
    if (status.indexOf('GAGAL') === 0) return { state: 'failed' };
    return { state: 'pending' };
  }
  return { state: 'notfound' };
}

/**
 * Jalankan SETELAH deploy Web App: memasang tautan halaman unduh
 * ke pesan konfirmasi Form. Boleh dijalankan ulang kapan saja.
 */
function updateFormConfirmation() {
  var url = ScriptApp.getService().getUrl();
  if (!url) throw new Error('Web App belum di-deploy. Deploy > New deployment > Web app dulu.');
  var formId = PropertiesService.getScriptProperties().getProperty('FORM_ID');
  FormApp.openById(formId).setConfirmationMessage(CONFIG.CONFIRMATION_TEXT + '\n\nUNDUH SERTIFIKAT: ' + url);
  Logger.log('Halaman unduh: ' + url);
}

var DOWNLOAD_PAGE_ = `<!doctype html><html><head><base target="_top"><style>
body{font-family:Arial,sans-serif;background:#fdf1ec;margin:0;padding:24px 16px;color:#1f2a44}
.card{max-width:440px;margin:0 auto;background:#fff;border-radius:12px;padding:24px;box-shadow:0 2px 10px rgba(0,0,0,.08)}
h1{font-size:20px;margin:0 0 6px}p{font-size:14px;line-height:1.5;margin:0 0 16px}
label{font-size:14px;font-weight:bold}
input{width:100%;box-sizing:border-box;padding:12px;margin:6px 0 14px;border:1px solid #bbb;border-radius:8px;font-size:16px}
button,a.btn{display:block;text-align:center;text-decoration:none;width:100%;box-sizing:border-box;padding:13px;border:0;border-radius:8px;background:#1b6b4a;color:#fff;font-size:16px;font-weight:bold;cursor:pointer}
button:disabled{opacity:.6}#msg{margin-top:16px;font-size:14px;line-height:1.5}.err{color:#b3261e}
</style></head><body><div class="card">
<h1>Unduh Sertifikat</h1>
<p>Webinar Dharma Wanita Persatuan Kabupaten Bengkalis Tahun 2026. Masukkan email yang Anda pakai saat mengisi daftar hadir.</p>
<label for="email">Email</label>
<input id="email" type="email" placeholder="nama@email.com" autocomplete="email">
<button id="go" onclick="cari()">Cari Sertifikat</button>
<div id="msg"></div></div>
<script>
var tries = 0, MAX = 12;
function el(id) { return document.getElementById(id); }
function cari() { tries = 0; el('go').disabled = true; el('msg').textContent = 'Mencari sertifikat...'; kirim(); }
function kirim() {
  google.script.run
    .withSuccessHandler(hasil)
    .withFailureHandler(function () { selesai('Terjadi gangguan. Coba lagi sebentar lagi.', true); })
    .findCertificate(el('email').value);
}
function hasil(r) {
  if (r.state === 'ready') {
    el('go').disabled = false;
    var msg = el('msg');
    msg.textContent = 'Sertifikat atas nama ' + r.name + ' sudah siap.';
    var a = document.createElement('a');
    a.className = 'btn'; a.href = r.url; a.textContent = 'Unduh PDF'; a.style.marginTop = '14px';
    msg.appendChild(a);
    return;
  }
  if (r.state === 'invalid') return selesai('Format email belum benar.', true);
  if (r.state === 'failed') return selesai('Sertifikat gagal dibuat. Hubungi panitia.', true);
  if (++tries < MAX) {
    el('msg').textContent = 'Sertifikat sedang dibuat, mohon tunggu...';
    setTimeout(kirim, 4000);
  } else {
    selesai(r.state === 'pending'
      ? 'Sertifikat masih diproses. Coba klik tombol lagi sebentar lagi.'
      : 'Email tidak ditemukan. Pastikan sama dengan email di form, atau hubungi panitia.', true);
  }
}
function selesai(t, isErr) { el('go').disabled = false; var m = el('msg'); m.textContent = t; m.className = isErr ? 'err' : ''; }
</script></body></html>`;

// ============================== HELPER ==============================
function getOutputFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* dibuat ulang di bawah */ }
  }
  var folder = DriveApp.createFolder(CONFIG.OUTPUT_FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

/** Pastikan kolom Status & Link ada, lalu kembalikan peta {judulKolom: nomorKolom}. */
function ensureColumns_(sheet) {
  var map = readHeaders_(sheet);
  [COL.STATUS, COL.LINK].forEach(function (title) {
    if (!map[title]) {
      var next = sheet.getLastColumn() + 1;
      sheet.getRange(1, next).setValue(title).setFontWeight('bold');
      map[title] = next;
    }
  });
  if (!map[COL.NAMA]) {
    throw new Error('Kolom "' + COL.NAMA + '" tidak ditemukan di Sheet. Cek judul pertanyaan Form.');
  }
  return map;
}

function readHeaders_(sheet) {
  var values = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  values.forEach(function (v, i) {
    var title = String(v).trim();
    if (title === 'Email Address' || title === 'Alamat Email') title = COL.EMAIL; // form berbahasa Indonesia/Inggris
    if (title) map[title] = i + 1;
  });
  return map;
}

function formatName_(raw) {
  var name = String(raw).replace(/\s+/g, ' ').trim();
  if (CONFIG.NAME_CASE === 'upper') return name.toUpperCase();
  if (CONFIG.NAME_CASE === 'title') {
    return name.toLowerCase().replace(/(^|[\s.\-'])([a-z])/g, function (m, p, c) { return p + c.toUpperCase(); });
  }
  return name;
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

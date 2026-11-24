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
    linkCell.setValue(pdfFile.getUrl());

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
    return folder.createFile(pdfBlob);
  } finally {
    copy.setTrashed(true); // buang salinan Slides sementara
  }
}

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

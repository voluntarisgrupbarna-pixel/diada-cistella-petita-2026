/**
 * 2a Diada Cistella Petita al Barna — Backend Apps Script
 *
 * CONFIGURACIÓ (abans de desplegar):
 *   1. Obre script.google.com → Nou projecte → enganxa aquest codi
 *   2. Crea un Google Sheet → copia l'ID de la URL → SHEET_ID
 *   3. Crea una carpeta a Drive "Comprovants Diada 2026" → copia l'ID → FOLDER_ID
 *   4. Desplegaments → Nou desplegament → Aplicació web
 *        Executar com: Jo (el meu compte)
 *        Qui pot accedir: Qualsevol (fins i tot anònim)
 *   5. Copia l'URL del desplegament
 *   6. Enganxa-la a la constant APPS_SCRIPT_URL de index.html
 *   7. Prova enviant el formulari i comprova el Sheet + l'email
 */

// ── CONFIGURACIÓ ──────────────────────────────────────────────────────────────
const SHEET_ID   = '1LFuJCTbQXA7jpZEcflapG6ZpIecjNtxskxeJdsVGk-U';
const FOLDER_ID  = '1OJDQaOfZmHulz8AznnwDcyNPLWG-MPuM';
const EMAIL_DEST = 'voluntarisgrupbarna@gmail.com';
// ─────────────────────────────────────────────────────────────────────────────

// GET: stats / list (JSONP) o redirecció a la landing
function doGet(e) {
  const action   = e && e.parameter && e.parameter.action;
  const callback = (e && e.parameter && e.parameter.callback) || 'handleData';

  if (action === 'stats') {
    const data = _getStats();
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  if (action === 'list') {
    const data = _getList();
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return HtmlService.createHtmlOutput(
    '<meta http-equiv="refresh" content="0;url=https://cistella.cbgrupbarna.info">' +
    '<p>Redirigint...</p>'
  );
}

// Retorna llista individual d'inscrits per a check-in / impressió
function _getList() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Inscripcions Diada');
  if (!sheet) return { inscrits: [], updatedAt: '' };

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[1]);

  const inscrits = rows.map((row, i) => ({
    num:       i + 1,
    nom:       (row[1] || '').toString().trim(),
    talla:     (row[4] || '—').toString().trim(),
    equip:     (row[5] || '—').toString().trim(),
    categoria: (row[6] || '—').toString().trim(),
    genere:    (row[7] || '—').toString().trim()
  }));

  // Ordenar: categoria → genere → nom
  const CAT_ORD = { 'Escoleta': 0, 'Premini': 1, 'Mini': 2 };
  inscrits.sort((a, b) => {
    const ca = CAT_ORD[a.categoria] ?? 9;
    const cb = CAT_ORD[b.categoria] ?? 9;
    if (ca !== cb) return ca - cb;
    if (a.genere !== b.genere) return a.genere.localeCompare(b.genere);
    return a.nom.localeCompare(b.nom);
  });

  return {
    inscrits,
    total: inscrits.length,
    updatedAt: Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm')
  };
}

function _getStats() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Inscripcions Diada');
  if (!sheet) return { total: 0, byGenere: {}, byCategoria: {}, byTalla: {}, equipsList: [], matrix: {} };

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[1]); // skip header + buits

  const byGenere          = {};
  const byCategoria       = {};
  const byTalla           = {};
  const byEquip           = {};
  const matrix            = {}; // { genere: { categoria: count } }

  rows.forEach(row => {
    const talla     = (row[4] || '—').toString().trim();
    const equip     = (row[5] || '—').toString().trim();
    const categoria = (row[6] || '—').toString().trim();
    const genere    = (row[7] || '—').toString().trim();

    byGenere[genere]        = (byGenere[genere]        || 0) + 1;
    byCategoria[categoria]  = (byCategoria[categoria]  || 0) + 1;
    byTalla[talla]          = (byTalla[talla]          || 0) + 1;
    byEquip[equip]          = (byEquip[equip]          || 0) + 1;

    if (!matrix[genere]) matrix[genere] = {};
    matrix[genere][categoria] = (matrix[genere][categoria] || 0) + 1;
  });

  const equipsList = Object.entries(byEquip)
    .map(([nom, count]) => ({ nom, count }))
    .sort((a, b) => b.count - a.count);

  return {
    total: rows.length,
    byGenere,
    byCategoria,
    byTalla,
    equipsList,
    matrix,
    updatedAt: Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm')
  };
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const timestamp = new Date().toLocaleString('ca-ES', { timeZone: 'Europe/Madrid' });

    // Pujar fitxer a Drive si n'hi ha
    let urlFitxer = '';
    if (data.fitxer && data.fitxer.dades) {
      urlFitxer = _guardarFitxer(data.fitxer);
    }

    // Guardar a Sheets
    _guardarAlSheet(timestamp, data, urlFitxer);

    // Enviar email de notificació
    _enviarEmail(timestamp, data, urlFitxer);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────────────────────────────────────

function _guardarFitxer(fitxer) {
  try {
    // El base64 ve com "data:tipus/ext;base64,XXXX..."
    const base64 = fitxer.dades.split(',')[1];
    const bytes  = Utilities.base64Decode(base64);
    const blob   = Utilities.newBlob(bytes, fitxer.tipus, fitxer.nom);

    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();
  } catch (err) {
    Logger.log('Error pujant fitxer: ' + err.message);
    return 'Error pujant fitxer';
  }
}

function _guardarAlSheet(timestamp, data, urlFitxer) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('Inscripcions Diada');

  if (!sheet) {
    sheet = ss.insertSheet('Inscripcions Diada');
    // Capçaleres
    sheet.appendRow([
      'Timestamp',
      'Nom menor',
      'Nom tutor/mare/pare',
      'Telèfon tutor',
      'Talla samarreta',
      'Equip',
      'Categoria',
      'Gènere',
      'Comprovant pagament (URL)'
    ]);
    // Format capçaleres
    const header = sheet.getRange(1, 1, 1, 9);
    header.setFontWeight('bold');
    header.setBackground('#CC0000');
    header.setFontColor('#ffffff');
  }

  sheet.appendRow([
    timestamp,
    data.nomMenor   || '',
    data.nomTutor   || '',
    data.telefon    || '',
    data.talla      || '',
    data.equip      || '',
    data.categoria  || '',
    data.genere     || '',
    urlFitxer
  ]);
}

function _enviarEmail(timestamp, data, urlFitxer) {
  const subject = `[Diada Cistella Petita] Nova inscripció: ${data.nomMenor}`;

  const body = `
Nova inscripció rebuda a la 2a Diada Cistella Petita al Barna
─────────────────────────────────────────────────────────────

📅 Data d'inscripció: ${timestamp}

👦 JUGADOR/A
   Nom i cognoms: ${data.nomMenor}
   Categoria: ${data.categoria}
   Equip: ${data.equip}
   Gènere: ${data.genere}
   Talla samarreta: ${data.talla}

👨‍👩‍👦 TUTOR / PARE / MARE
   Nom i cognoms: ${data.nomTutor}
   Telèfon: ${data.telefon}

💳 COMPROVANT DE PAGAMENT
   ${urlFitxer || 'No adjuntat'}

─────────────────────────────────────────────────────────────
Consulta totes les inscripcions al Google Sheet:
https://docs.google.com/spreadsheets/d/${SHEET_ID}
`;

  GmailApp.sendEmail(EMAIL_DEST, subject, body);
}

/**
 * Funció de prova — executa manualment des de l'editor per comprovar
 * que tot funciona sense necessitat d'enviar el formulari real.
 */
function _testInscripcio() {
  const dadesTost = {
    nomMenor:  'Marc Puig Fernández',
    nomTutor:  'Joan Puig Martínez',
    telefon:   '612000000',
    talla:     '8 anys',
    equip:     'Premini Masculí A',
    categoria: 'Premini',
    genere:    'Masculí',
    fitxer:    null
  };
  _guardarAlSheet(new Date().toLocaleString('ca-ES'), dadesTost, 'TEST - sense fitxer');
  _enviarEmail(new Date().toLocaleString('ca-ES'), dadesTost, 'TEST - sense fitxer');
  Logger.log('Test completat ✓');
}

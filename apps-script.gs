/**
 * Observation tennis de table — réception des bilans dans Google Sheets
 *
 * Installation (une seule fois, par le professeur) :
 * 1. Créer une Google Sheet vide.
 * 2. Menu Extensions > Apps Script, effacer le contenu et coller ce fichier.
 * 3. Enregistrer, puis Déployer > Nouveau déploiement > type « Application Web ».
 *    - Exécuter en tant que : Moi
 *    - Qui a accès : Tout le monde
 * 4. Autoriser l'accès, puis copier l'URL qui se termine par /exec.
 * 5. Coller cette URL dans index.html : const SHEET_URL = '...';
 */

const HEADERS = [
  'id', 'date', 'classe', 'observateur', 'observe', 'numero', 'mode', 'total', 'pct_table', 'pct_fautes',
  'adv_z1', 'adv_z2', 'adv_z3', 'adv_z4', 'adv_z5', 'adv_z6',
  'eleve_z1', 'eleve_z2', 'eleve_z3', 'eleve_z4', 'eleve_z5', 'eleve_z6',
  'filet', 'hors_table', 'critere', 'reussites', 'atteint', 'pct_vise', 'pct_evite', 'commentaire', 'recu_le'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const now = new Date();
    (data.rows || []).forEach(function (r) {
      r.recu_le = now;
      sheet.appendRow(HEADERS.map(function (h) { return r[h] !== undefined ? r[h] : ''; }));
    });
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Observation tennis de table : réception active.');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Bilans');
  if (!sheet) sheet = ss.insertSheet('Bilans');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

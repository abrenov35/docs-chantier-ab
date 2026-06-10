// ── Configuration ────────────────────────────────────────────────────────────
// Définir via Fichier > Propriétés du projet > Propriétés du script :
//   SPREADSHEET_ID  → l'ID de la Google Sheet
//   ANTHROPIC_KEY   → votre clé API Anthropic
const SPREADSHEET_ID    = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const ANTHROPIC_KEY     = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
const SHEET_NAME        = 'Documents';
const ANTHROPIC_URL     = 'https://api.anthropic.com/v1/messages';
const MODEL             = 'claude-opus-4-8';
const DRIVE_FOLDER_NAME = 'docs-chantier-photos';

// Photos dont le base64 dépasse ce seuil (~500 KB décodés) sont ignorées
const MAX_PHOTO_CHARS = 700000;

// Colonnes de la feuille (1-based)
// A: id | B: type | C: chantier | D: operateur | E: date
// F: contenu | G: history | H: (réservé) | I: photoUrls
const COL = { id:1, type:2, chantier:3, operateur:4, date:5, contenu:6, history:7, photoUrls:8 };

// ── Point d'entrée ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const action = p.action;

    if (action === 'list')         return json(listDocs());
    if (action === 'get')          return json(getDoc(p.id));
    if (action === 'delete')       return json(deleteDoc(p.id));
    if (action === 'save' || action === 'update') return json(saveOrUpdate(p));
    if (action === 'upload_photo') return json(uploadOnePhoto(p));

    if (p.modif) return json(callClaude({ modif: p.modif, history: p.history }));

    return json(callClaude({ payload: p }));
  } catch(err) {
    return json({ ok: false, error: err.message });
  }
}

function doGet() {
  return ContentService.createTextOutput('OK');
}

// ── Drive ─────────────────────────────────────────────────────────────────────
function getOrCreateFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

// Upload d'une seule photo — appelé par l'action "upload_photo".
// Reçoit { docId, photoIndex, base64, mimeType, filename }.
// Stocke l'URL dans col 9 à l'index photoIndex pour conserver l'ordre.
function uploadOnePhoto(p) {
  const { docId, photoIndex, base64, mimeType, filename } = p;

  if (!base64 || base64.length > MAX_PHOTO_CHARS) {
    Logger.log('Photo ignorée [' + photoIndex + '] — trop grande (' + (base64 || '').length + ' chars)');
    return { ok: false, error: 'Photo trop grande' };
  }

  const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  const ext    = ((filename || '').split('.').pop() || 'jpg').toLowerCase();
  const mime   = { jpg:'image/jpeg', jpeg:'image/jpeg', png:'image/png',
                   gif:'image/gif',  webp:'image/webp', heic:'image/heic' }[ext]
                 || mimeType || 'image/jpeg';
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), mime,
                                   filename || ('photo_' + photoIndex + '.jpg'));
  const file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url    = 'https://drive.google.com/uc?export=view&id=' + file.getId();

  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL.id - 1] === docId) {
      let urls = [];
      try { urls = JSON.parse(rows[i][COL.photoUrls - 1] || '[]'); } catch(e) {}
      urls[photoIndex] = url;
      sheet.getRange(i + 1, COL.photoUrls).setValue(JSON.stringify(urls));
      Logger.log('upload_photo OK [' + photoIndex + ']: ' + url);
      return { ok: true, url };
    }
  }
  return { ok: false, error: 'Document introuvable : ' + docId };
}

// ── CRUD Sheets ───────────────────────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

// Sauvegarde texte + history + photos (URLs Dropbox en colonne 9).
function saveOrUpdate(p) {
  const sheet = getSheet();
  const { action, id, type, chantier, operateur, date, contenu, history, files } = p;
  const historyJson   = JSON.stringify(history || []);
  const photoUrlsJson = files && files.length ? JSON.stringify(files.map(f => ({ name: f.name, type: f.type, url: f.url }))) : null;

  if (action === 'update' && id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][COL.id - 1] === id) {
        const r = i + 1;
        sheet.getRange(r, COL.type).setValue(type);
        sheet.getRange(r, COL.chantier).setValue(chantier);
        sheet.getRange(r, COL.operateur).setValue(operateur);
        sheet.getRange(r, COL.date).setValue(date);
        sheet.getRange(r, COL.contenu).setValue(contenu);
        sheet.getRange(r, COL.history).setValue(historyJson);
        if (photoUrlsJson !== null) sheet.getRange(r, COL.photoUrls).setValue(photoUrlsJson);
        return { id };
      }
    }
  }

  const newId = generateId();
  const row   = new Array(COL.photoUrls).fill('');
  row[COL.id        - 1] = newId;
  row[COL.type      - 1] = type;
  row[COL.chantier  - 1] = chantier;
  row[COL.operateur - 1] = operateur;
  row[COL.date      - 1] = date;
  row[COL.contenu   - 1] = contenu;
  row[COL.history   - 1] = historyJson;
  if (photoUrlsJson !== null) row[COL.photoUrls - 1] = photoUrlsJson;
  sheet.appendRow(row);
  return { id: newId };
}

function getDoc(id) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[COL.id - 1] !== id) continue;

    let history = [];
    try { history = JSON.parse(r[COL.history - 1] || '[]'); } catch(e) {}

    let photoUrls = [];
    try { photoUrls = JSON.parse(r[COL.photoUrls - 1] || '[]'); } catch(e) {}

    const files = photoUrls
      .filter(item => item)
      .map((item, idx) => ({
        name: (typeof item === 'object' ? item.name : null) || ('photo_' + (idx + 1) + '.jpg'),
        type: (typeof item === 'object' ? item.type : null) || 'image/jpeg',
        url:  (typeof item === 'object' ? item.url  : item) || null
      }));

    return {
      id:        r[COL.id        - 1],
      type:      r[COL.type      - 1],
      chantier:  r[COL.chantier  - 1],
      operateur: r[COL.operateur - 1],
      date:      r[COL.date      - 1],
      contenu:   r[COL.contenu   - 1],
      history,
      files
    };
  }
  throw new Error('Document introuvable : ' + id);
}

function listDocs() {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  const docs  = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[COL.id - 1]) docs.push({
      id:        r[COL.id        - 1],
      type:      r[COL.type      - 1],
      chantier:  r[COL.chantier  - 1],
      operateur: r[COL.operateur - 1],
      date:      r[COL.date      - 1]
    });
  }
  return { docs: docs.reverse() };
}

function deleteDoc(id) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL.id - 1] === id) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

// ── Claude API ────────────────────────────────────────────────────────────────
function callClaude({ payload, modif, history }) {
  let messages;

  if (modif) {
    messages = [...(history || []), { role: 'user', content: modif }];
  } else {
    const { type, chantier, operateur, observations, files } = payload;
    const labels = { CRCh:'Compte Rendu de Chantier', FCh:'Fiche Chantier', PVR:'Procès-Verbal de Réception' };
    const userContent = [{ type: 'text', text: 'Observations :\n' + observations }];
    // Seuls les fichiers non-image (PDF, docs) peuvent encore être joints ici
    (files || []).filter(f => f.data).forEach(f => {
      if (f.type && f.type.startsWith('image/')) return; // les photos arrivent via upload_photo
      userContent.push({ type: 'text', text: '[Fichier joint : ' + f.name + ']' });
    });
    messages = [{ role: 'user', content: userContent }];

    const system = [
      'Tu es un assistant expert en documents de chantier BTP pour AB RENOV 35 (Rennes).',
      'Génère un ' + (labels[type] || type) + ' professionnel structuré.',
      'Chantier : ' + chantier + (operateur ? ' | Opérateur : ' + operateur : '') + ' | Date : ' + new Date().toLocaleDateString('fr-FR'),
      'Structure : sections en ## , sous-sections en ### , listes à puces avec - .',
      "Ne répète pas le titre ni les infos d'en-tête dans le corps du document."
    ].join('\n');

    const resp  = anthropicFetch(system, messages);
    const text  = resp.content.filter(b => b.type === 'text').map(b => b.text).join('');
    return { text, history: [...messages, { role: 'assistant', content: resp.content }] };
  }

  const resp  = anthropicFetch(null, messages);
  const text  = resp.content.filter(b => b.type === 'text').map(b => b.text).join('');
  return { text, history: [...messages, { role: 'assistant', content: resp.content }] };
}

function anthropicFetch(system, messages) {
  const body = { model: MODEL, max_tokens: 4096, messages };
  if (system) body.system = system;
  const resp = UrlFetchApp.fetch(ANTHROPIC_URL, {
    method: 'POST',
    contentType: 'application/json',
    headers: { 'x-api-key': ANTHROPIC_KEY, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  const data = JSON.parse(resp.getContentText());
  if (data.error) throw new Error(data.error.message);
  return data;
}

// ── Helper ────────────────────────────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

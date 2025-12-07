const SPREADSHEET_ID = '1w46H58534iN35C55oZHbs4jUpNc6IGX1_NME5ASVbhE';
const INVENTORY_SHEET = 'Inventario';
const LOGBOOK_SHEET = 'Bitacora';
const REQUESTS_SHEET = 'Solicitudes';
const PHOTOS_FOLDER_NAME = 'Inventario Fotos';

function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Gestión de Laboratorio')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheetByName(name) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(name);
  return sheet || ss.insertSheet(name);
}

function ensureHeaders(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = existing.some((v, i) => !v && headers[i]);
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getInventory() {
  const sheet = getSheetByName(INVENTORY_SHEET);
  const headers = ['Nombre', 'Cantidad', 'Estado', 'Categoría', 'Descripción', 'Foto', 'Última actualización'];
  ensureHeaders(sheet, headers);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return data.map((row, idx) => ({
    row: idx + 2,
    nombre: row[0],
    cantidad: row[1],
    estado: row[2],
    categoria: row[3],
    descripcion: row[4],
    foto: row[5],
    actualizado: row[6]
  }));
}

function saveInventoryItem(payload) {
  const sheet = getSheetByName(INVENTORY_SHEET);
  const headers = ['Nombre', 'Cantidad', 'Estado', 'Categoría', 'Descripción', 'Foto', 'Última actualización'];
  ensureHeaders(sheet, headers);
  const photoUrl = payload.photoData ? storePhoto(payload.photoData, payload.nombre) : payload.foto || '';
  const values = [
    payload.nombre,
    Number(payload.cantidad) || 0,
    payload.estado,
    payload.categoria,
    payload.descripcion || '',
    photoUrl,
    new Date()
  ];
  if (payload.row) {
    sheet.getRange(Number(payload.row), 1, 1, values.length).setValues([values]);
    return { message: 'Elemento actualizado', photoUrl: photoUrl };
  }
  sheet.appendRow(values);
  return { message: 'Elemento agregado', photoUrl: photoUrl };
}

function storePhoto(photoDataUrl, name) {
  const match = photoDataUrl.match(/^data:(.*?);base64,(.*)$/);
  if (!match) return '';
  const contentType = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const blob = Utilities.newBlob(bytes, contentType, `${name || 'foto'}-${Date.now()}.png`);
  const folder = getOrCreateFolder(PHOTOS_FOLDER_NAME);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function addLogEntry(entry) {
  const sheet = getSheetByName(LOGBOOK_SHEET);
  const headers = ['Fecha', 'Preparador', 'Actividad', 'Elementos', 'Solicitante'];
  ensureHeaders(sheet, headers);
  const values = [
    entry.fecha ? new Date(entry.fecha) : new Date(),
    entry.preparador || 'N/D',
    entry.actividad || '',
    entry.elementos || '',
    entry.solicitante || ''
  ];
  sheet.appendRow(values);
  return 'Bitácora actualizada';
}

function getLogEntries() {
  const sheet = getSheetByName(LOGBOOK_SHEET);
  const headers = ['Fecha', 'Preparador', 'Actividad', 'Elementos', 'Solicitante'];
  ensureHeaders(sheet, headers);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return data.map((row, idx) => ({
    row: idx + 2,
    fecha: row[0],
    preparador: row[1],
    actividad: row[2],
    elementos: row[3],
    solicitante: row[4]
  }));
}

function submitRequest(request) {
  const sheet = getSheetByName(REQUESTS_SHEET);
  const headers = ['Fecha de solicitud', 'Docente', 'Inicio', 'Fin', 'Elementos', 'Extras', 'Estado', 'Foto entrega', 'Preparador'];
  ensureHeaders(sheet, headers);
  const values = [
    new Date(),
    request.docente,
    request.inicio ? new Date(request.inicio) : '',
    request.fin ? new Date(request.fin) : '',
    request.elementos || '',
    request.extras || '',
    'Pendiente',
    '',
    ''
  ];
  sheet.appendRow(values);
  return 'Solicitud enviada';
}

function getRequests() {
  const sheet = getSheetByName(REQUESTS_SHEET);
  const headers = ['Fecha de solicitud', 'Docente', 'Inicio', 'Fin', 'Elementos', 'Extras', 'Estado', 'Foto entrega', 'Preparador'];
  ensureHeaders(sheet, headers);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return data.map((row, idx) => ({
    row: idx + 2,
    fechaSolicitud: row[0],
    docente: row[1],
    inicio: row[2],
    fin: row[3],
    elementos: row[4],
    extras: row[5],
    estado: row[6],
    fotoEntrega: row[7],
    preparador: row[8]
  }));
}

function markRequestPrepared(payload) {
  const sheet = getSheetByName(REQUESTS_SHEET);
  const headers = ['Fecha de solicitud', 'Docente', 'Inicio', 'Fin', 'Elementos', 'Extras', 'Estado', 'Foto entrega', 'Preparador'];
  ensureHeaders(sheet, headers);
  const row = Number(payload.row);
  const photoUrl = payload.photoData ? storePhoto(payload.photoData, `Solicitud-${row}`) : payload.fotoEntrega || '';
  sheet.getRange(row, 7).setValue('Preparado');
  sheet.getRange(row, 8).setValue(photoUrl);
  sheet.getRange(row, 9).setValue(payload.preparador || '');
  return 'Solicitud marcada como preparada';
}

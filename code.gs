const SPREADSHEET_ID = '1w46H58534iN35C55oZHbs4jUpNc6IGX1_NME5ASVbhE';
const BITACORA_SPREADSHEET_ID = '19rvs-kBt9o87d40-8nIFUtv8_KnKXnxPfwegUT9h24A';
const INVENTORY_SHEET = 'Inventario';
const LOGBOOK_SHEET = 'Bitacora';
const REQUESTS_SHEET = 'Solicitudes';
const USERS_SHEET = 'Usuarios';
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

function getBitacoraSpreadsheet() {
  return SpreadsheetApp.openById(BITACORA_SPREADSHEET_ID);
}

function getSheetByName(name) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(name);
  return sheet || ss.insertSheet(name);
}

function getBitacoraSheet(name) {
  const ss = getBitacoraSpreadsheet();
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

function ensureDefaultUsers() {
  const sheet = getSheetByName(USERS_SHEET);
  const headers = ['Email', 'Nombre', 'Rol'];
  ensureHeaders(sheet, headers);
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), headers.length).getValues();
  const hasPreparador = data.some(r => (r[2] || '').toLowerCase() === 'preparador');
  const hasDocente = data.some(r => (r[2] || '').toLowerCase() === 'docente');
  const rowsToAdd = [];
  if (!hasPreparador) {
    rowsToAdd.push(['preparador@demo.com', 'Preparador Demo', 'Preparador']);
  }
  if (!hasDocente) {
    rowsToAdd.push(['docente@demo.com', 'Docente Demo', 'Docente']);
  }
  if (rowsToAdd.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
  }
}

function login(email) {
  if (!email) {
    throw new Error('Ingresa un correo para continuar');
  }
  const sheet = getSheetByName(USERS_SHEET);
  ensureHeaders(sheet, ['Email', 'Nombre', 'Rol']);
  ensureDefaultUsers();
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 3).getValues();
  const user = data.find(r => (r[0] || '').toLowerCase() === email.toLowerCase());
  if (!user) {
    throw new Error('Usuario no encontrado. Solicita al preparador que te dé de alta.');
  }
  return { email: user[0], nombre: user[1], rol: user[2] };
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
  const sheet = getBitacoraSheet(LOGBOOK_SHEET);
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
  const sheet = getBitacoraSheet(LOGBOOK_SHEET);
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

function importLogbookCsv(csvText) {
  if (!csvText) {
    throw new Error('El archivo CSV está vacío');
  }
  const rows = Utilities.parseCsv(csvText);
  if (!rows.length) {
    throw new Error('No se encontraron datos en el CSV');
  }
  const sheet = getBitacoraSheet(LOGBOOK_SHEET);
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return 'Bitácora actualizada desde CSV';
}

function exportLogbookCsv() {
  const sheet = getBitacoraSheet(LOGBOOK_SHEET);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) {
    return '';
  }
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  return data.map(row => row.map(value => value === null ? '' : value).join(',')).join('\n');
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
  addLogEntry({
    fecha: new Date(),
    preparador: 'Docente',
    actividad: `Solicitud de práctica (${request.docente || 'N/D'})`,
    elementos: request.elementos,
    solicitante: request.docente
  });
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
  addLogEntry({
    fecha: new Date(),
    preparador: payload.preparador || 'Preparador',
    actividad: 'Solicitud preparada',
    elementos: sheet.getRange(row, 5).getValue(),
    solicitante: sheet.getRange(row, 2).getValue()
  });
  return 'Solicitud marcada como preparada';
}

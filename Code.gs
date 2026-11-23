/**
 * CRM Admisiones, Familias y Entrevistas (Modelo REA)
 * Backend gratuito en Google Sheets + Google Apps Script
 *
 * INSTRUCCIONES:
 * 1. Crea una Hoja de Cálculo nueva en Google Drive.
 * 2. Ve a Extensiones -> Apps Script, borra el contenido y pega este archivo.
 * 3. Pulsa Implementar -> Nueva implementación -> Tipo: "Aplicación web".
 *    - Ejecutar como: "Yo"
 *    - Quién tiene acceso: "Cualquier usuario"
 * 4. Copia la URL /exec y pégala en la pestaña Administración del CRM.
 */

const JSON_SHEET = '_CRM_JSON';

function doGet(e) {
  try {
    const state = readStateFromSheet_();
    return jsonResponse_({ ok: true, state });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'saveState' && body.state) {
      writeStateToSheet_(body.state);
      syncReadableTabs_(body.state);
      return jsonResponse_({ ok: true, savedAt: new Date().toISOString() });
    }
    return jsonResponse_({ ok: false, error: 'Acción no reconocida' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function readStateFromSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(JSON_SHEET);
  if (!sh) return null;
  const values = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues();
  const raw = values.map(r => r[0]).join('');
  if (!raw) return null;
  return JSON.parse(raw);
}

function writeStateToSheet_(state) {
  const sh = getOrCreateSheet_(JSON_SHEET);
  try { sh.hideSheet(); } catch (_) {}
  const raw = JSON.stringify(state);
  const CHUNK_SIZE = 45000;
  const rows = [];
  for (let i = 0; i < raw.length; i += CHUNK_SIZE) {
    rows.push([raw.slice(i, i + CHUNK_SIZE)]);
  }
  sh.clearContents();
  if (rows.length > 0) {
    sh.getRange(1, 1, rows.length, 1).setValues(rows);
  }
}

function syncReadableTabs_(state) {
  // 1. Pestaña Familias
  const famSh = getOrCreateSheet_('Familias');
  famSh.clearContents();
  const famHeaders = [
    'ID', 'Familia', 'Progenitor 1', 'Progenitor 2', 'Teléfono', 'Email',
    'Alumno', 'Edad', 'Etapa', 'Curso Objetivo', 'Colegio Origen',
    'Etapa Embudo', 'Probabilidad %', 'Origen', 'Responsable',
    'Intereses / Necesidades', 'Entrevistador', 'Fecha Entrevista', 'Acta Entrevista', 'Actualizado'
  ];
  const famRows = (state.leads || []).map(l => [
    l.id || '', l.familyName || '', l.parent1 || '', l.parent2 || '', l.phone || '', l.email || '',
    l.studentName || '', l.studentAge || '', l.stageEdu || '', l.targetGrade || '', l.originSchool || '',
    l.pipelineStage || '', l.probability || 0, l.origin || '', l.owner || '',
    l.interests || '', l.interviewer || '', l.interviewDate || '', l.interviewNotes || '', l.updatedAt || ''
  ]);
  famSh.getRange(1, 1, 1, famHeaders.length).setValues([famHeaders]).setFontWeight('bold');
  if (famRows.length) famSh.getRange(2, 1, famRows.length, famHeaders.length).setValues(famRows);

  // 2. Pestaña Visitas
  const visSh = getOrCreateSheet_('Visitas');
  visSh.clearContents();
  const visHeaders = ['ID', 'ID Familia', 'Tipo', 'Responsable', 'Fecha', 'Hora', 'Duración (min)', 'Lugar', 'Notas'];
  const visRows = (state.visits || []).map(v => [
    v.id || '', v.leadId || '', v.type || '', v.staff || '', v.date || '', v.time || '', v.duration || 45, v.location || '', v.notes || ''
  ]);
  visSh.getRange(1, 1, 1, visHeaders.length).setValues([visHeaders]).setFontWeight('bold');
  if (visRows.length) visSh.getRange(2, 1, visRows.length, visHeaders.length).setValues(visRows);

  // 3. Pestaña Equipo
  const teamSh = getOrCreateSheet_('Equipo');
  teamSh.clearContents();
  const teamHeaders = ['ID', 'Nombre', 'Rol / Cargo', 'Etapas', 'Entrevistador'];
  const teamRows = (state.team || []).map(m => [
    m.id || '', m.name || '', m.role || '', m.stages || '', m.interviewer ? 'Sí' : 'No'
  ]);
  teamSh.getRange(1, 1, 1, teamHeaders.length).setValues([teamHeaders]).setFontWeight('bold');
  if (teamRows.length) teamSh.getRange(2, 1, teamRows.length, teamHeaders.length).setValues(teamRows);
}

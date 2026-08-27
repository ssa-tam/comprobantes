// ===========================================================
//  SST — Sistema de respaldo de información
//  Google Apps Script — Web App  v9
//  + Control de acceso por unidad (habilitado/deshabilitado)
//  + Panel de administrador
// ===============================================================

const SHEET_ID      = "1PyjT6E8FOiaeH79eKCMKIOiQbpmKqcpoyCfVr7Sd9ww";
const CONTROL_ID    = "1b3EczybnJhSsJwvz156T1bJQOzytyt1zRTAw1RX6IP4";
const SHEET_NAME    = "Accesos";
const CONTROL_NAME  = "Control_Acceso";
const DATA_START    = 4;
const CACHE_KEY     = "sst_idx_v9";
const CONTROL_CACHE = "sst_ctrl_v9";
const CACHE_TTL     = 21600;

const ADMIN_USER = "ADMIN_SST";
const STATUS_SHEET = "Status_Trimestres";
const ADMIN_PASS = "Admin2024*";

const QNA_ORDER = ["QNA 01","QNA 02","QNA 03","QNA 04","QNA 05","QNA 06",
                   "QNA 07","QNA 08","QNA 09","QNA 10","QNA 11","QNA 12",
                   "QNA 13","QNA 14","QNA 15","QNA 16","QNA 17","QNA 18"];

const TRIM_ORDER = ["1° Trimestre","2° Trimestre","3° Trimestre","4° Trimestre"];

function extractFolderId(value) {
  if (!value) return "";
  const str = value.toString().trim();
  if (!str.includes("/")) return str;
  const match = str.match(/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : str;
}

// ── Índice de accesos ───────────────────────────────────────
function buildIndex() {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const raw     = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, 10).getValues();
  const index   = {};

  for (const row of raw) {
    const trimestre = (row[1] || "").toString().trim();
    const unidad    = (row[2] || "").toString().trim();
    const user      = (row[3] || "").toString().trim();
    const pass      = (row[4] || "").toString().trim();
    const qna       = (row[5] || "").toString().trim();
    const nomina    = (row[6] || "").toString().trim();
    const tipo      = (row[7] || "").toString().trim();
    const folder    = extractFolderId((row[8] || "").toString().trim());

    if (!user || !folder) continue;

    const key = user + "|" + pass;
    if (!index[key]) index[key] = { unidad, opciones: [] };
    index[key].opciones.push({ trimestre, qna, nomina, tipo, folderId: folder });
  }
  return index;
}

function getIndex() {
  const cache = CacheService.getScriptCache();
  const meta  = cache.get(CACHE_KEY + "_meta");
  if (meta) {
    const { numChunks } = JSON.parse(meta);
    let json = "";
    for (let i = 0; i < numChunks; i++) {
      const chunk = cache.get(CACHE_KEY + "_" + i);
      if (!chunk) return saveIndex();
      json += chunk;
    }
    return JSON.parse(json);
  }
  return saveIndex();
}

function saveIndex() {
  const index = buildIndex();
  const json  = JSON.stringify(index);
  const cache = CacheService.getScriptCache();
  const CHUNK = 90000;
  const chunks = Math.ceil(json.length / CHUNK);
  for (let i = 0; i < chunks; i++) {
    cache.put(CACHE_KEY + "_" + i, json.slice(i * CHUNK, (i + 1) * CHUNK), CACHE_TTL);
  }
  cache.put(CACHE_KEY + "_meta", JSON.stringify({ numChunks: chunks }), CACHE_TTL);
  return index;
}

// ── Control de acceso ────────────────────────────────────────
function getControlMap() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get(CONTROL_CACHE);
  if (cached) return JSON.parse(cached);

  const sheet   = SpreadsheetApp.openById(CONTROL_ID).getSheetByName(CONTROL_NAME);
  const lastRow = sheet.getLastRow();
  const raw     = sheet.getRange(4, 1, lastRow - 3, 4).getValues();

  const map = {};
  for (const row of raw) {
    const user   = (row[2] || "").toString().trim();
    const estado = (row[3] || "").toString().trim();
    if (user) map[user] = estado;
  }

  cache.put(CONTROL_CACHE, JSON.stringify(map), 300);
  return map;
}

function isHabilitado(usuario) {
  const map    = getControlMap();
  const estado = map[usuario];
  if (!estado) return true;
  return estado.toLowerCase() === "activo";
}

// ── Admin: leer lista de unidades ────────────────────────────
function getUnidadesList() {
  const sheet   = SpreadsheetApp.openById(CONTROL_ID).getSheetByName(CONTROL_NAME);
  const lastRow = sheet.getLastRow();
  const raw     = sheet.getRange(4, 1, lastRow - 3, 4).getValues();
  return raw.map((row, i) => ({
    num:    row[0],
    nombre: (row[1] || "").toString().trim(),
    usuario:(row[2] || "").toString().trim(),
    estado: (row[3] || "").toString().trim(),
    fila:   i + 4
  })).filter(u => u.usuario);
}

// ── Admin: actualizar estado ─────────────────────────────────
function setEstado(usuario, nuevoEstado) {
  const sheet   = SpreadsheetApp.openById(CONTROL_ID).getSheetByName(CONTROL_NAME);
  const lastRow = sheet.getLastRow();
  const raw     = sheet.getRange(4, 3, lastRow - 3, 2).getValues();

  for (let i = 0; i < raw.length; i++) {
    const rowUser = (raw[i][0] || "").toString().trim();
    if (rowUser === usuario) {
      sheet.getRange(i + 4, 4).setValue(nuevoEstado);
      break;
    }
  }
  CacheService.getScriptCache().remove(CONTROL_CACHE);
}

function setTodas(nuevoEstado) {
  const sheet   = SpreadsheetApp.openById(CONTROL_ID).getSheetByName(CONTROL_NAME);
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - 3;
  const values  = Array(numRows).fill([nuevoEstado]);
  sheet.getRange(4, 4, numRows, 1).setValues(values);
  CacheService.getScriptCache().remove(CONTROL_CACHE);
}

// ── Limpiar caché ────────────────────────────────────────────
function limpiarCache() {
  const cache = CacheService.getScriptCache();
  const meta  = cache.get(CACHE_KEY + "_meta");
  if (meta) {
    const { numChunks } = JSON.parse(meta);
    for (let i = 0; i < numChunks; i++) cache.remove(CACHE_KEY + "_" + i);
  }
  cache.remove(CACHE_KEY + "_meta");
  cache.remove(CONTROL_CACHE);
  Logger.log("Cache limpiado");
}

// ── doGet ────────────────────────────────────────────────────
function doGet(e) {
  const usuario  = (e.parameter.usuario   || "").trim();
  const password = (e.parameter.password  || "").trim();
  const callback = (e.parameter.callback  || "").trim();
  const action   = (e.parameter.action    || "").trim();
  const step     = (e.parameter.step      || "1").trim();
  const trimSel  = (e.parameter.trimestre || "").trim();
  const qnaSel   = (e.parameter.qna       || "").trim();

  if (!usuario && !password && !action) {
    return respond({ status: "SST Web App activa v9" }, callback);
  }

  // ── MARCAR TRIMESTRE COMPLETADO (unidad) ─────────────────
  if (action === "marcar_completado" && usuario && password) {
    try {
      const index2  = getIndex();
      const key2    = usuario + "|" + password;
      const entry2  = index2[key2];
      if (!entry2) return respond({ ok: false, msg: "credenciales_invalidas" }, callback);
      const trimMark = (e.parameter.trimestre || "").toString().trim();
      Logger.log("Marcando: " + usuario + " | " + trimMark);
      if (!trimMark) return respond({ ok: false, msg: "trimestre_vacio" }, callback);
      setEstatus_Trimestre(usuario, entry2.unidad, trimMark, "Completado");
      return respond({ ok: true, msg: "marcado", trimestre: trimMark }, callback);
    } catch(err) {
      return respond({ ok: false, msg: "error_servidor", detail: err.message }, callback);
    }
  }

  // ── ACCIONES DE ADMIN ────────────────────────────────────
  if (action) {
    if (usuario !== ADMIN_USER || password !== ADMIN_PASS) {
      return respond({ ok: false, msg: "no_autorizado" }, callback);
    }
    if (action === "get_unidades") {
      return respond({ ok: true, unidades: getUnidadesList() }, callback);
    }
    if (action === "set_estado") {
      const target   = (e.parameter.target || "").trim();
      const newState = (e.parameter.estado || "").trim();
      if (!target || !newState) return respond({ ok: false, msg: "parametros_faltantes" }, callback);
      setEstado(target, newState);
      return respond({ ok: true, msg: "actualizado" }, callback);
    }
    if (action === "set_todas") {
      const newState = (e.parameter.estado || "").trim();
      if (!newState) return respond({ ok: false, msg: "parametros_faltantes" }, callback);
      setTodas(newState);
      return respond({ ok: true, msg: "todas_actualizadas" }, callback);
    }
    if (action === "get_all_estatus") {
      return respond({ ok: true, estatus: getAllEstatus() }, callback);
    }
    if (action === "reset_estatus") {
      const target    = (e.parameter.target    || "").trim();
      const trimestre = (e.parameter.trimestre || "").trim();
      const obs       = (e.parameter.obs       || "").trim();
      if (!target || !trimestre) return respond({ ok: false, msg: "parametros_faltantes" }, callback);
      setEstatus_Trimestre(target, "", trimestre, "Pendiente", obs);
      return respond({ ok: true, msg: "estatus_reseteado" }, callback);
    }
    if (action === "marcar_visto") {
      const target    = (e.parameter.target    || "").trim();
      const trimestre = (e.parameter.trimestre || "").trim();
      if (!target || !trimestre) return respond({ ok: false, msg: "parametros_faltantes" }, callback);
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(STATUS_SHEET);
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
        for (let i = 0; i < data.length; i++) {
          if ((data[i][0] || "").toString().trim() === target &&
              (data[i][2] || "").toString().trim() === trimestre) {
            sheet.getRange(i + 2, 7).setValue("Sí");
            break;
          }
        }
      }
      return respond({ ok: true, msg: "visto_marcado" }, callback);
    }
    return respond({ ok: false, msg: "accion_desconocida" }, callback);
  }

  // ── LOGIN DE UNIDAD ──────────────────────────────────────
  if (!usuario || !password) {
    return respond({ ok: false, msg: "campos_vacios" }, callback);
  }

  try {
    if (!isHabilitado(usuario)) {
      return respond({ ok: false, msg: "deshabilitado" }, callback);
    }

    const index   = getIndex();
    const key     = usuario + "|" + password;
    const entry   = index[key];

    if (!entry) return respond({ ok: false, msg: "credenciales_invalidas" }, callback);

    const unidad  = entry.unidad;
    const opciones = entry.opciones;

    if (opciones.length === 0) return respond({ ok: false, msg: "sin_carpeta" }, callback);

    // PASO 3: filtrar por trimestre + QNA
    if (step === "3" && trimSel && qnaSel) {
      const filtered = opciones.filter(o => o.trimestre === trimSel && o.qna === qnaSel);
      const ordenadas = ordenarOpciones(filtered);
      return respond({ ok: true, step: 3, unidad, trimestre: trimSel, qna: qnaSel, opciones: ordenadas }, callback);
    }

    // PASO 2: QNAs de un trimestre
    if (step === "2" && trimSel) {
      const trimOpciones = opciones.filter(o => o.trimestre === trimSel);
      const qnas = [];
      QNA_ORDER.forEach(q => {
        const count = trimOpciones.filter(o => o.qna === q).length;
        if (count > 0) qnas.push({ qna: q, count });
      });
      if (qnas.length === 1) {
        const filtered = trimOpciones.filter(o => o.qna === qnas[0].qna);
        return respond({ ok: true, step: 3, unidad, trimestre: trimSel, qna: qnas[0].qna, opciones: filtered }, callback);
      }
      return respond({ ok: true, step: 2, unidad, trimestre: trimSel, qnas }, callback);
    }

    // PASO 1: trimestres disponibles
    const trimestres = [];
    TRIM_ORDER.forEach(t => {
      const count = opciones.filter(o => o.trimestre === t).length;
      if (count > 0) trimestres.push({ trimestre: t, count });
    });

    if (trimestres.length === 1) {
      const t = trimestres[0].trimestre;
      const trimOpciones = opciones.filter(o => o.trimestre === t);
      const qnas = [];
      QNA_ORDER.forEach(q => {
        const count = trimOpciones.filter(o => o.qna === q).length;
        if (count > 0) qnas.push({ qna: q, count });
      });
      return respond({ ok: true, step: 2, unidad, trimestre: t, qnas, skippedStep1: true }, callback);
    }

    return respond({ ok: true, step: 1, unidad, trimestres }, callback);

  } catch (err) {
    return respond({ ok: false, msg: "error_servidor", detail: err.message }, callback);
  }
}


// ── Ordenar opciones: Especial → Ord → Extras ───────────────
function ordenarOpciones(opciones) {
  // Nóminas especiales en orden deseado
  const ESPECIALES = [
    "2da Aguinaldo",
    "Reyes",
    "Bono Día Madres",
    "Bono Día Padres",
    "Prima Vacacional"
  ];
  // Nóminas normales en orden deseado
  const NORMALES = ["Ord", "Extras"];

  function getPrioridad(nomina) {
    const iEsp = ESPECIALES.indexOf(nomina);
    if (iEsp >= 0) return iEsp; // 0..4
    const iNor = NORMALES.indexOf(nomina);
    if (iNor >= 0) return ESPECIALES.length + iNor; // 5, 6
    return 99;
  }

  // Orden de tipos dentro de cada nómina
  const TIPO_ORDER = ["411","610","Estatal","Eventual","For 1","For 2","For 3","Homologados","IMSS-Bienestar","Jubilados","Reg N","Reg S","Aportación EST"];

  function getTipoPrioridad(tipo) {
    const i = TIPO_ORDER.indexOf(tipo);
    return i >= 0 ? i : 99;
  }

  return opciones.slice().sort((a, b) => {
    const pA = getPrioridad(a.nomina);
    const pB = getPrioridad(b.nomina);
    if (pA !== pB) return pA - pB;
    // Mismo grupo de nómina: ordenar por tipo
    return getTipoPrioridad(a.tipo) - getTipoPrioridad(b.tipo);
  });
}


// ── Estatus de trimestre ─────────────────────────────────────
function getEstatus(usuario, trimestre) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(STATUS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { completado: false, fecha: "", obs: "", visto: true };
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (const row of data) {
    if ((row[0] || "").toString().trim() === usuario &&
        (row[2] || "").toString().trim() === trimestre) {
      return {
        completado: (row[3] || "").toString().trim() === "Completado",
        fecha:      (row[4] || "").toString().trim(),
        obs:        (row[5] || "").toString().trim(),
        visto:      (row[6] || "").toString().trim() === "Sí"
      };
    }
  }
  return { completado: false, fecha: "", obs: "", visto: true };
}

function setEstatus_Trimestre(usuario, unidad, trimestre, nuevoEstatus, observaciones) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(STATUS_SHEET);
  const lastRow = sheet.getLastRow();
  const fecha = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy HH:mm");
  const obs = observaciones || "";

  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      if ((data[i][0] || "").toString().trim() === usuario &&
          (data[i][2] || "").toString().trim() === trimestre) {
        sheet.getRange(i + 2, 4).setValue(nuevoEstatus);
        sheet.getRange(i + 2, 5).setValue(fecha);
        if (nuevoEstatus === "Pendiente") {
          sheet.getRange(i + 2, 6).setValue(obs);  // observaciones
          sheet.getRange(i + 2, 7).setValue("");    // limpiar visto
        } else {
          sheet.getRange(i + 2, 6).setValue("");    // limpiar obs al completar
          sheet.getRange(i + 2, 7).setValue("");
        }
        return;
      }
    }
  }
  // No existe — crear fila nueva
  sheet.appendRow([usuario, unidad, trimestre, nuevoEstatus, fecha, obs, ""]);
}

function getAllEstatus() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(STATUS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  return data.map(row => ({
    usuario:   (row[0] || "").toString().trim(),
    unidad:    (row[1] || "").toString().trim(),
    trimestre: (row[2] || "").toString().trim(),
    estatus:   (row[3] || "").toString().trim(),
    fecha:     (row[4] || "").toString().trim(),
    obs:       (row[5] || "").toString().trim(),
    visto:     (row[6] || "").toString().trim() === "Sí"
  })).filter(r => r.usuario);
}

function respond(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
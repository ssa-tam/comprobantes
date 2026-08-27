// ============================================================
//  SST — Sistema de respaldo de información
//  Google Apps Script — Web App  v7
//  Con CacheService para acelerar lecturas repetidas del Sheets
// =============================================================

const SHEET_ID   = "1LrpGDZBLk_Ko29783_wDYhwWs-d521tD9TEBjvDDaPo"; // SST_Completo_1erTrimestre
const SHEET_NAME = "Accesos";
const DATA_START = 4;
const CACHE_KEY  = "sst_accesos_data_v1";
const CACHE_TTL_SECONDS = 21600; // 6 horas

// Columnas (1-indexed): A=1 #, B=2 Unidad, C=3 Usuario, D=4 Contraseña,
// E=5 QNA, F=6 Nómina, G=7 Tipo, H=8 ID Carpeta, I=9 Estado

function extractFolderId(value) {
  if (!value) return "";
  const str = value.toString().trim();
  if (!str.includes("/")) return str;
  const match = str.match(/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : str;
}

// Devuelve un array de filas "compactas" [unidad, usuario, pass, qna, nomina, tipo, folderId]
// Solo incluye filas que SÍ tienen folderId (las vacías no sirven para el login)
function getCompactData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);

  if (cached) {
    return JSON.parse(cached);
  }

  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const raw     = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, 9).getValues();

  const compact = [];
  for (const row of raw) {
    const unidad = (row[1] || "").toString().trim();
    const user   = (row[2] || "").toString().trim();
    const pass   = (row[3] || "").toString().trim();
    const qna    = (row[4] || "").toString().trim();
    const nomina = (row[5] || "").toString().trim();
    const tipo   = (row[6] || "").toString().trim();
    const folder = extractFolderId((row[7] || "").toString().trim());

    if (!user || !folder) continue; // omitir filas sin usuario o sin carpeta asignada

    compact.push([unidad, user, pass, qna, nomina, tipo, folder]);
  }

  // CacheService tiene límite de 100KB por entrada -> dividir en chunks si es necesario
  const json = JSON.stringify(compact);
  const CHUNK_SIZE = 90000; // ~90KB por chunk, deja margen

  if (json.length <= CHUNK_SIZE) {
    cache.put(CACHE_KEY, json, CACHE_TTL_SECONDS);
  } else {
    // Guardar en chunks
    const numChunks = Math.ceil(json.length / CHUNK_SIZE);
    const chunkMap = {};
    for (let i = 0; i < numChunks; i++) {
      const chunk = json.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
      cache.put(CACHE_KEY + "_chunk_" + i, chunk, CACHE_TTL_SECONDS);
    }
    cache.put(CACHE_KEY + "_meta", JSON.stringify({ numChunks: numChunks }), CACHE_TTL_SECONDS);
  }

  return compact;
}

// Lee considerando posibles chunks
function getCompactDataWithChunks() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const meta = cache.get(CACHE_KEY + "_meta");
  if (meta) {
    const { numChunks } = JSON.parse(meta);
    let json = "";
    for (let i = 0; i < numChunks; i++) {
      const chunk = cache.get(CACHE_KEY + "_chunk_" + i);
      if (!chunk) return getCompactData(); // cache incompleto -> recargar
      json += chunk;
    }
    return JSON.parse(json);
  }

  return getCompactData();
}

function doGet(e) {
  const usuario  = (e.parameter.usuario  || "").trim();
  const password = (e.parameter.password || "").trim();
  const callback = (e.parameter.callback || "").trim();
  const step     = (e.parameter.step     || "1").trim();
  const qnaSel   = (e.parameter.qna      || "").trim();

  if (!usuario && !password) {
    return respond({ status: "SST Web App activa v7 (con cache)" }, callback);
  }

  if (!usuario || !password) {
    return respond({ ok: false, msg: "campos_vacios" }, callback);
  }

  try {
    const data = getCompactDataWithChunks();

    let unidadNombre = null;
    const opciones = [];

    for (const row of data) {
      const [rowUnidad, rowUser, rowPass, rowQna, rowNomina, rowTipo, rowFolder] = row;

      if (rowUser === usuario && rowPass === password) {
        unidadNombre = rowUnidad;
        opciones.push({
          qna: rowQna,
          nomina: rowNomina,
          tipo: rowTipo,
          folderId: rowFolder
        });
      }
    }

    if (!unidadNombre) {
      return respond({ ok: false, msg: "credenciales_invalidas" }, callback);
    }

    if (opciones.length === 0) {
      return respond({ ok: false, msg: "sin_carpeta" }, callback);
    }

    // PASO 2: ya se sabe la QNA, devolver solo esas opciones
    if (step === "2" && qnaSel) {
      const filtered = opciones.filter(o => o.qna === qnaSel);
      return respond({
        ok: true,
        step: 2,
        unidad: unidadNombre,
        qna: qnaSel,
        opciones: filtered
      }, callback);
    }

    // Si solo hay UNA opción en total, abrir directo
    if (opciones.length === 1) {
      return respond({
        ok: true,
        single: true,
        unidad: unidadNombre,
        folderId: opciones[0].folderId
      }, callback);
    }

    // PASO 1: conteos por QNA
    const qnaOrder = ["QNA 01","QNA 02","QNA 03","QNA 04","QNA 05","QNA 06"];
    const qnasDisponibles = [];
    qnaOrder.forEach(q => {
      const count = opciones.filter(o => o.qna === q).length;
      if (count > 0) {
        qnasDisponibles.push({ qna: q, count: count });
      }
    });

    // Si solo hay UNA QNA con opciones, saltar directo al paso 2
    if (qnasDisponibles.length === 1) {
      const onlyQna = qnasDisponibles[0].qna;
      const filtered = opciones.filter(o => o.qna === onlyQna);
      return respond({
        ok: true,
        step: 2,
        unidad: unidadNombre,
        qna: onlyQna,
        opciones: filtered,
        skippedStep1: true
      }, callback);
    }

    return respond({
      ok: true,
      step: 1,
      unidad: unidadNombre,
      single: false,
      qnas: qnasDisponibles
    }, callback);

  } catch (err) {
    return respond({ ok: false, msg: "error_servidor", detail: err.message }, callback);
  }
}

// Función para forzar refresco de cache manualmente (ejecutar desde el editor
// si actualizas el Sheets y quieres que el cambio se refleje antes de 6 horas)
function limpiarCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_KEY);
  cache.remove(CACHE_KEY + "_meta");
  for (let i = 0; i < 10; i++) {
    cache.remove(CACHE_KEY + "_chunk_" + i);
  }
  Logger.log("Cache limpiado");
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
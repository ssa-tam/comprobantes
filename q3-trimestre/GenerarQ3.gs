// ===============================================================
//  SST — Generar 3er Trimestre en su propio Sheet independiente
//  Corre UNA SOLA VEZ desde el editor de Apps Script (función
//  generarTercerTrimestre) después de pegar este código junto con
//  Codigo_Q3.gs en el mismo proyecto de Apps Script.
//
//  Qué hace:
//   1. Lee las credenciales (usuario/contraseña) de cada unidad
//      desde Control_Acceso (misma fuente que usa el sistema
//      principal), para no tener que capturarlas de nuevo.
//   2. Camina la carpeta "3° TRIMESTRE" en Drive (QNA → Nómina →
//      Tipo → Unidad, y también carpetas especiales/bonos que
//      cuelgan directo de la QNA) y arma una fila por cada carpeta
//      de unidad que encuentra.
//   3. Escribe esas filas en el Sheet SST_3er_Trimestre → pestaña
//      "Accesos", con el mismo formato que usa el sistema actual.
// ===============================================================

const SHEET_ID_Q3    = "1C7d_klfgAa4tBxjXhb8hnOLHhrGr6cVsBpoDLuHVOZM"; // SST_3er_Trimestre
const SHEET_NAME_Q3  = "Accesos";
const CONTROL_ID_Q3  = "1b3EczybnJhSsJwvz156T1bJQOzytyt1zRTAw1RX6IP4"; // Control_Acceso (compartido)
const CONTROL_NAME_Q3 = "Control_Acceso";
const FOLDER_Q3       = "1FCzGamb_dU3ntUmYU8o6DMeUXbbPlqyZ"; // 3° TRIMESTRE
const TRIMESTRE_Q3    = "3° Trimestre";

// ── Carpetas especiales (bonos) de este trimestre ──────────────
// Edita este objeto ANTES de correr el script: la llave es la QNA
// donde vive la carpeta especial, el valor es el nombre EXACTO de
// la carpeta tal como aparece en Drive (ej. "3. Bono Día Niño").
// Déjalo vacío {} si este trimestre no tiene carpetas especiales.
const BONOS_Q3 = {
  // "QNA 15": "Bono Día Niño",
  // "QNA 18": "Aguinaldo"
};

function generarTercerTrimestre() {
  // 1) Construir mapa unidad -> {usuario, pass} desde Control_Acceso.
  //    Control_Acceso solo trae usuario/estado, no contraseña — la
  //    contraseña se toma de cualquier fila existente del sistema
  //    principal (SST_Todos_Trimestres) para esa unidad.
  const unitMap = construirMapaCredenciales_();
  Logger.log("Unidades con credenciales: " + Object.keys(unitMap).length);

  const sheet = abrirOCrearHojaAccesos_();
  const lastRow = sheet.getLastRow();
  const rows = [];
  let contador = 0;

  const folderQ3 = DriveApp.getFolderById(FOLDER_Q3);
  const qnaFolders = folderQ3.getFolders();

  while (qnaFolders.hasNext()) {
    const qnaFolder = qnaFolders.next();
    const qnaLabel = normalizarQNA_(qnaFolder.getName().trim());
    Logger.log("Procesando " + qnaLabel);

    const nominaFolders = qnaFolder.getFolders();
    while (nominaFolders.hasNext()) {
      const nominaFolder = nominaFolders.next();
      const nominaNameRaw = nominaFolder.getName().trim(); // "1. Ord", "2. Extras", "3. Bono..."
      const nominaLabel = nominaNameRaw.replace(/^\d+\.\s*/, "").trim();

      const esBono = Object.values(BONOS_Q3).indexOf(nominaLabel) >= 0 ||
                     Object.values(BONOS_Q3).indexOf(nominaNameRaw) >= 0;

      if (esBono) {
        // Carpeta especial: sin nivel de "tipo", va directo a unidades
        contador += agregarFilasDeUnidades_(rows, nominaFolder, qnaLabel, nominaLabel, "", unitMap);
      } else {
        // Nómina normal (Ord/Extras): tiene nivel de "tipo" antes de unidad
        const tipoFolders = nominaFolder.getFolders();
        while (tipoFolders.hasNext()) {
          const tipoFolder = tipoFolders.next();
          const tipoName = tipoFolder.getName().trim();
          contador += agregarFilasDeUnidades_(rows, tipoFolder, qnaLabel, nominaLabel, tipoName, unitMap);
        }
      }
    }
  }

  if (rows.length === 0) {
    Logger.log("❌ No se encontraron filas. Revisa la estructura de carpetas o BONOS_Q3.");
    return;
  }

  const startRow = lastRow + 1;
  sheet.getRange(startRow, 1, rows.length, 10).setValues(rows);

  Logger.log("✅ Filas agregadas: " + contador);
  Logger.log("✅ Desde fila " + startRow + " hasta " + (startRow + rows.length - 1));
  Logger.log("Recuerda: después de correr esto, limpia el caché del Web App (limpiarCache) si ya está desplegado.");
}

// ── Helpers ──────────────────────────────────────────────────

function agregarFilasDeUnidades_(rows, contenedorFolder, qnaLabel, nominaLabel, tipoName, unitMap) {
  let n = 0;
  const unidadFolders = contenedorFolder.getFolders();
  while (unidadFolders.hasNext()) {
    const unidadFolder = unidadFolders.next();
    const unidadName = unidadFolder.getName().trim();
    const folderId = unidadFolder.getId();
    const creds = unitMap[unidadName];
    if (!creds) {
      Logger.log("⚠️ Sin credenciales para: " + unidadName + " (agrégala en Control_Acceso o en el Sheet principal)");
      continue;
    }
    rows.push([
      "", TRIMESTRE_Q3, unidadName, creds.usuario, creds.pass,
      qnaLabel, nominaLabel, tipoName, folderId, "✓ Asignada"
    ]);
    n++;
  }
  return n;
}

function construirMapaCredenciales_() {
  // Contraseñas: se toman de las filas ya existentes en el Sheet
  // principal (mismo criterio que usaba Llenado_2do_Trimestre.gs).
  const sheetPrincipal = SpreadsheetApp.openById("1PyjT6E8FOiaeH79eKCMKIOiQbpmKqcpoyCfVr7Sd9ww").getSheetByName("Accesos");
  const lastRow = sheetPrincipal.getLastRow();
  const raw = sheetPrincipal.getRange(4, 1, lastRow - 3, 10).getValues();

  const map = {};
  for (const row of raw) {
    const unidad = (row[2] || "").toString().trim();
    const usuario = (row[3] || "").toString().trim();
    const pass = (row[4] || "").toString().trim();
    if (unidad && usuario && !map[unidad]) {
      map[unidad] = { usuario, pass };
    }
  }
  return map;
}

function abrirOCrearHojaAccesos_() {
  const ss = SpreadsheetApp.openById(SHEET_ID_Q3);
  let sheet = ss.getSheetByName(SHEET_NAME_Q3);
  if (!sheet) {
    // El Sheet se creó a partir de un CSV, la pestaña quedó con otro
    // nombre (ej. "Untitled") — la renombramos a "Accesos".
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME_Q3);
  }
  return sheet;
}

function normalizarQNA_(nombre) {
  const match = nombre.match(/(\d+)/);
  if (match) {
    const num = parseInt(match[1]);
    return "QNA " + (num < 10 ? "0" + num : num);
  }
  return nombre;
}

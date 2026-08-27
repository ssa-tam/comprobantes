// ===============================================================
//  SST / Comprobantes Digitales SSA — Panel de Administrador
//  Google Apps Script — Web App independiente (doGet sirve HTML)
//
//  Qué resuelve:
//   - Un dashboard donde Susy ve los 4 trimestres (Q1-Q4) y abre
//     cualquiera con un clic.
//   - Al abrir un trimestre que TODAVÍA NO tiene su carpeta/Sheet
//     generados, primero pregunta qué carpetas especiales (bonos)
//     aplican ese trimestre, y con esa respuesta genera automático
//     toda la estructura de carpetas (clonando el patrón de un
//     trimestre ya armado) + llena su Sheet de accesos.
//   - Si el trimestre ya existe, simplemente muestra su estado y
//     un link directo a su carpeta y a su Sheet.
//
//  Despliegue: Implementar > Nueva implementación > Tipo: App web
//              Ejecutar como: Yo | Acceso: Solo tú (o quien tú
//              decidas) — este panel es solo para administración,
//              a diferencia del Web App de las unidades.
// ===============================================================

const CFG_SHEET_ID   = "1b3EczybnJhSsJwvz156T1bJQOzytyt1zRTAw1RX6IP4"; // usamos Control_Acceso como base
const CFG_SHEET_NAME = "Config_Trimestres"; // se crea sola la primera vez
const ADMIN_USER_DASH = "ADMIN_SST";
const ADMIN_PASS_DASH = "Admin2024*"; // ⚠ cámbiala antes de compartir el link

const ANO_2026_FOLDER_ID = "1Cb6_hC1sdnQCP_ebdvcAxu0dIvrtEAJe"; // carpeta "AÑO 2026"

// Lista sugerida de carpetas especiales (bonos). El panel también
// deja escribir un nombre libre si no está en esta lista.
const BONOS_SUGERIDOS = ["2da Aguinaldo", "Reyes", "Bono Día Madres", "Bono Día Padres", "Prima Vacacional"];

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("AdminPanel")
    .setTitle("Comprobantes Digitales SSA — Admin")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ── Auth simple para las llamadas desde el panel ────────────────
function login_(usuario, password) {
  if (usuario !== ADMIN_USER_DASH || password !== ADMIN_PASS_DASH) {
    throw new Error("no_autorizado");
  }
}

// ── Config de trimestres (se guarda en su propia pestaña) ──────
function getConfigSheet_() {
  const ss = SpreadsheetApp.openById(CFG_SHEET_ID);
  let sheet = ss.getSheetByName(CFG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CFG_SHEET_NAME);
    sheet.appendRow(["Trimestre", "FolderId", "SheetId", "DeployUrl", "Estado", "Bonos"]);
    ["1° Trimestre", "2° Trimestre", "3° Trimestre", "4° Trimestre"].forEach(t => {
      sheet.appendRow([t, "", "", "", "Pendiente", ""]);
    });
    // Precarga lo que ya sabemos que existe
    setConfigRow_("1° Trimestre", { folderId: "1iUycLSuXVN1rrEBSNmawDnUtKSQo762F", estado: "Generado" });
    setConfigRow_("2° Trimestre", { folderId: "1FPoHf9p_OM-XoxxE37BJTd67A2eQK_V8", sheetId: "1PyjT6E8FOiaeH79eKCMKIOiQbpmKqcpoyCfVr7Sd9ww", estado: "Generado" });
    setConfigRow_("3° Trimestre", { folderId: "1FCzGamb_dU3ntUmYU8o6DMeUXbbPlqyZ", sheetId: "1C7d_klfgAa4tBxjXhb8hnOLHhrGr6cVsBpoDLuHVOZM", estado: "Generado" });
  }
  return sheet;
}

function setConfigRow_(trimestre, patch) {
  const sheet = SpreadsheetApp.openById(CFG_SHEET_ID).getSheetByName(CFG_SHEET_NAME) || getConfigSheet_();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === trimestre) {
      const row = i + 2;
      if (patch.folderId !== undefined) sheet.getRange(row, 2).setValue(patch.folderId);
      if (patch.sheetId !== undefined) sheet.getRange(row, 3).setValue(patch.sheetId);
      if (patch.deployUrl !== undefined) sheet.getRange(row, 4).setValue(patch.deployUrl);
      if (patch.estado !== undefined) sheet.getRange(row, 5).setValue(patch.estado);
      if (patch.bonos !== undefined) sheet.getRange(row, 6).setValue(patch.bonos);
      return;
    }
  }
}

// ── Llamado por el panel: usuario, password, y el resto de acciones ──
function panelListarTrimestres(usuario, password) {
  login_(usuario, password);
  const sheet = getConfigSheet_();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return data.map(r => ({
    trimestre: r[0], folderId: r[1], sheetId: r[2], deployUrl: r[3], estado: r[4] || "Pendiente", bonos: r[5]
  }));
}

// Devuelve las carpetas especiales sugeridas para el picker
function panelGetBonosSugeridos(usuario, password) {
  login_(usuario, password);
  return BONOS_SUGERIDOS;
}

// ── Generar un trimestre nuevo: clona estructura + llena Sheet ─
// bonosSeleccionados: array de nombres de carpetas especiales que
// el admin marcó en el panel (puede venir vacío).
function panelGenerarTrimestre(usuario, password, trimestre, bonosSeleccionados) {
  login_(usuario, password);

  const numero = { "1° Trimestre": 1, "2° Trimestre": 2, "3° Trimestre": 3, "4° Trimestre": 4 }[trimestre];
  if (!numero) throw new Error("trimestre_invalido");

  const cfgSheet = getConfigSheet_();
  const cfg = panelListarTrimestres(usuario, password).find(t => t.trimestre === trimestre);
  if (cfg && cfg.estado === "Generado") {
    return { ok: true, msg: "ya_existe", folderId: cfg.folderId, sheetId: cfg.sheetId };
  }

  // 1) Carpeta del trimestre dentro de "AÑO 2026"
  const anoFolder = DriveApp.getFolderById(ANO_2026_FOLDER_ID);
  const nombreTrimestre = numero + "° TRIMESTRE";
  const existentes = anoFolder.getFoldersByName(nombreTrimestre);
  const folderTrimestre = existentes.hasNext() ? existentes.next() : anoFolder.createFolder(nombreTrimestre);

  // 2) Trimestre plantilla del que clonamos la estructura (el más
  //    reciente que ya esté generado, para copiar Ord/Extras/tipos/unidades)
  const trimestres = panelListarTrimestres(usuario, password);
  const plantilla = trimestres.filter(t => t.estado === "Generado" && t.folderId).sort((a, b) =>
    ({ "1° Trimestre": 1, "2° Trimestre": 2, "3° Trimestre": 3, "4° Trimestre": 4 }[b.trimestre] -
     { "1° Trimestre": 1, "2° Trimestre": 2, "3° Trimestre": 3, "4° Trimestre": 4 }[a.trimestre]))[0];
  if (!plantilla) throw new Error("sin_plantilla_disponible");
  const folderPlantilla = DriveApp.getFolderById(plantilla.folderId);
  const qnaPlantilla = primeraSubcarpeta_(folderPlantilla); // ej. "QNA 07"
  if (!qnaPlantilla) throw new Error("plantilla_vacia");

  // 3) Crear las 6 QNAs de este trimestre clonando Ord/Extras/tipos/unidades
  //    (vacío, sin archivos) desde la QNA plantilla, más las carpetas
  //    especiales seleccionadas.
  const qnaInicio = (numero - 1) * 6 + 1;
  const nombresUnidades = listarUnidades_();
  const bonos = bonosSeleccionados || [];

  for (let i = 0; i < 6; i++) {
    const numQna = qnaInicio + i;
    const nombreQna = "QNA " + (numQna < 10 ? "0" + numQna : numQna);
    const existentesQna = folderTrimestre.getFoldersByName(nombreQna);
    const qnaFolder = existentesQna.hasNext() ? existentesQna.next() : folderTrimestre.createFolder(nombreQna);

    clonarNominaSiNoExiste_(qnaFolder, qnaPlantilla, "1. Ord");
    clonarNominaSiNoExiste_(qnaFolder, qnaPlantilla, "2. Extras");

    // Carpetas especiales: sin nivel de tipo, solo unidades
    bonos.forEach((nombreBono, idx) => {
      const numeroCarpeta = 3 + idx;
      const nombreCarpeta = numeroCarpeta + ". " + nombreBono;
      if (qnaFolder.getFoldersByName(nombreCarpeta).hasNext()) return;
      const bonoFolder = qnaFolder.createFolder(nombreCarpeta);
      nombresUnidades.forEach(u => bonoFolder.createFolder(u));
    });
  }

  // 4) Sheet independiente para este trimestre (se copia el layout del Sheet Q3)
  const nuevoSheet = SpreadsheetApp.create("SST_" + numeroATexto_(numero) + "_Trimestre");
  const sh = nuevoSheet.getSheets()[0];
  sh.setName("Accesos");
  sh.getRange(1, 1).setValue("SST — Accesos " + trimestre + " (Independiente)");
  sh.getRange(2, 1).setValue("⚠ Columna I = ID de carpeta Drive. Generado automáticamente desde el panel de admin.");
  sh.getRange(3, 1, 1, 10).setValues([["#", "Trimestre", "Unidad", "Usuario", "Contraseña", "QNA", "Nómina", "Tipo", "ID Carpeta Drive", "Estado"]]);
  DriveApp.getFileById(nuevoSheet.getId()).moveTo(DriveApp.getFolderById(getComprobantesFolderId_()));

  // 5) Llenar el Sheet caminando la carpeta recién creada
  const filas = generarFilasDesdeCarpeta_(folderTrimestre, trimestre, bonos);
  if (filas.length > 0) {
    sh.getRange(4, 1, filas.length, 10).setValues(filas);
  }

  // 6) Guardar en la config
  setConfigRow_(trimestre, {
    folderId: folderTrimestre.getId(),
    sheetId: nuevoSheet.getId(),
    estado: "Generado",
    bonos: bonos.join(", ")
  });

  return { ok: true, msg: "generado", folderId: folderTrimestre.getId(), sheetId: nuevoSheet.getId(), filas: filas.length };
}

// ── Helpers de clonado ──────────────────────────────────────────

function primeraSubcarpeta_(folder) {
  const it = folder.getFolders();
  return it.hasNext() ? it.next() : null;
}

function clonarNominaSiNoExiste_(qnaFolderDestino, qnaFolderOrigen, nombreNomina) {
  if (qnaFolderDestino.getFoldersByName(nombreNomina).hasNext()) return; // ya existe, no duplicar
  const origenIt = qnaFolderOrigen.getFoldersByName(nombreNomina);
  if (!origenIt.hasNext()) return;
  const origenNomina = origenIt.next();
  const destinoNomina = qnaFolderDestino.createFolder(nombreNomina);

  const tipoFolders = origenNomina.getFolders();
  while (tipoFolders.hasNext()) {
    const tipoOrigen = tipoFolders.next();
    const tipoDestino = destinoNomina.createFolder(tipoOrigen.getName());
    const unidadFolders = tipoOrigen.getFolders();
    while (unidadFolders.hasNext()) {
      tipoDestino.createFolder(unidadFolders.next().getName());
    }
  }
}

function listarUnidades_() {
  const sheet = SpreadsheetApp.openById(CFG_SHEET_ID).getSheetByName("Control_Acceso");
  const lastRow = sheet.getLastRow();
  const raw = sheet.getRange(4, 2, lastRow - 3, 1).getValues(); // columna B = nombre de unidad
  return raw.map(r => (r[0] || "").toString().trim()).filter(Boolean);
}

function generarFilasDesdeCarpeta_(folderTrimestre, trimestreLabel, bonos) {
  const unitMap = construirMapaCredencialesGlobal_();
  const rows = [];

  const qnaFolders = folderTrimestre.getFolders();
  while (qnaFolders.hasNext()) {
    const qnaFolder = qnaFolders.next();
    const qnaLabel = normalizarQNA2_(qnaFolder.getName());

    const nominaFolders = qnaFolder.getFolders();
    while (nominaFolders.hasNext()) {
      const nominaFolder = nominaFolders.next();
      const nominaNameRaw = nominaFolder.getName().trim();
      const nominaLabel = nominaNameRaw.replace(/^\d+\.\s*/, "").trim();
      const esBono = bonos.indexOf(nominaLabel) >= 0;

      if (esBono) {
        agregarFilas_(rows, nominaFolder, trimestreLabel, qnaLabel, nominaLabel, "", unitMap);
      } else {
        const tipoFolders = nominaFolder.getFolders();
        while (tipoFolders.hasNext()) {
          const tipoFolder = tipoFolders.next();
          agregarFilas_(rows, tipoFolder, trimestreLabel, qnaLabel, nominaLabel, tipoFolder.getName().trim(), unitMap);
        }
      }
    }
  }
  return rows;
}

function agregarFilas_(rows, contenedor, trimestreLabel, qnaLabel, nominaLabel, tipoName, unitMap) {
  const unidadFolders = contenedor.getFolders();
  while (unidadFolders.hasNext()) {
    const unidadFolder = unidadFolders.next();
    const unidadName = unidadFolder.getName().trim();
    const creds = unitMap[unidadName];
    if (!creds) continue; // unidad sin usuario conocido — se omite, revisar Control_Acceso
    rows.push(["", trimestreLabel, unidadName, creds.usuario, creds.pass, qnaLabel, nominaLabel, tipoName, unidadFolder.getId(), "✓ Asignada"]);
  }
}

function construirMapaCredencialesGlobal_() {
  const sheetPrincipal = SpreadsheetApp.openById("1PyjT6E8FOiaeH79eKCMKIOiQbpmKqcpoyCfVr7Sd9ww").getSheetByName("Accesos");
  const lastRow = sheetPrincipal.getLastRow();
  const raw = sheetPrincipal.getRange(4, 1, lastRow - 3, 10).getValues();
  const map = {};
  for (const row of raw) {
    const unidad = (row[2] || "").toString().trim();
    const usuario = (row[3] || "").toString().trim();
    const pass = (row[4] || "").toString().trim();
    if (unidad && usuario && !map[unidad]) map[unidad] = { usuario, pass };
  }
  return map;
}

function normalizarQNA2_(nombre) {
  const match = nombre.match(/(\d+)/);
  if (match) {
    const num = parseInt(match[1]);
    return "QNA " + (num < 10 ? "0" + num : num);
  }
  return nombre;
}

function numeroATexto_(n) {
  return ["", "1er", "2do", "3er", "4to"][n];
}

function getComprobantesFolderId_() {
  return "1FqxQBp9rmhX5rmFON8UWibj-W92aUgjD"; // carpeta "Comprobantes Digitales SSA"
}

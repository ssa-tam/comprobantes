// ==========================================================
//  SST — Llenar 2° Trimestre en SST_Todos_Trimestres
//  Corre UNA SOLA VEZ desde el editor de Apps Script
// ===============================================================

const SHEET_ID_2T   = "1PyjT6E8FOiaeH79eKCMKIOiQbpmKqcpoyCfVr7Sd9ww"; // SST_Todos_Trimestres
const SHEET_NAME_2T = "Accesos";
const FOLDER_2T     = "1FPoHf9p_OM-XoxxE37BJTd67A2eQK_V8"; // 2° TRIMESTRE
const PASSWORD_2T   = "SST2024*"; // misma contraseña del 1° Trimestre
const TRIMESTRE_2T  = "2° Trimestre";

// QNAs con sus nóminas especiales
const BONOS = {
  "QNA 09": "Bono Día Madres",
  "QNA 11": "Bono Día Padres",
  "QNA 12": "Prima Vacacional"
};

function llenar2doTrimestre() {
  const sheet    = SpreadsheetApp.openById(SHEET_ID_2T).getSheetByName(SHEET_NAME_2T);
  const lastRow  = sheet.getLastRow();
  const rows     = [];

  // Leer el Sheet para obtener usuarios por unidad (de filas existentes)
  const existing = sheet.getRange(4, 1, lastRow - 3, 10).getValues();
  
  // Construir mapa unidad -> { usuario, password }
  const unitMap = {};
  for (const row of existing) {
    const unidad  = (row[2] || "").toString().trim();
    const usuario = (row[3] || "").toString().trim();
    const pass    = (row[4] || "").toString().trim();
    if (unidad && usuario && !unitMap[unidad]) {
      unitMap[unidad] = { usuario, pass };
    }
  }

  Logger.log("Unidades encontradas: " + Object.keys(unitMap).length);

  // Abrir carpeta 2° Trimestre
  const folder2T = DriveApp.getFolderById(FOLDER_2T);
  const qnaFolders = folder2T.getFolders();

  let contador = 0;

  while (qnaFolders.hasNext()) {
    const qnaFolder = qnaFolders.next();
    const qnaName   = qnaFolder.getName().trim(); // ej. "QNA 07"
    const qnaLabel  = normalizarQNA(qnaName);

    Logger.log("Procesando: " + qnaLabel);

    // Iterar nóminas dentro de la QNA
    const nominaFolders = qnaFolder.getFolders();
    while (nominaFolders.hasNext()) {
      const nominaFolder = nominaFolders.next();
      const nominaName   = nominaFolder.getName().trim(); // ej. "1. Ord", "2. Extras"
      const nominaLabel  = extraerNomina(nominaName);

      // Iterar tipos dentro de la nómina
      const tipoFolders = nominaFolder.getFolders();
      while (tipoFolders.hasNext()) {
        const tipoFolder = tipoFolders.next();
        const tipoName   = tipoFolder.getName().trim(); // ej. "IMSS-Bienestar"

        // Iterar carpetas de unidades
        const unidadFolders = tipoFolder.getFolders();
        while (unidadFolders.hasNext()) {
          const unidadFolder = unidadFolders.next();
          const unidadName   = unidadFolder.getName().trim();
          const folderId     = unidadFolder.getId();

          // Buscar usuario de esta unidad
          const creds = unitMap[unidadName];
          if (!creds) {
            Logger.log("⚠️ Sin credenciales para: " + unidadName);
            continue;
          }

          rows.push([
            "",              // col A: # (vacío, se puede llenar después)
            TRIMESTRE_2T,    // col B: Trimestre
            unidadName,      // col C: Unidad
            creds.usuario,   // col D: Usuario
            creds.pass,      // col E: Contraseña
            qnaLabel,        // col F: QNA
            nominaLabel,     // col G: Nómina
            tipoName,        // col H: Tipo
            folderId,        // col I: ID Carpeta Drive
            "✓ Asignada"     // col J: Estado
          ]);
          contador++;
        }
      }

      // Agregar bono especial si aplica para esta QNA
      if (BONOS[qnaLabel]) {
        const bonoNombre = BONOS[qnaLabel];
        // Buscar si ya existe carpeta de bono en esta nómina
        // (el bono va como tipo adicional dentro de Ord y Extras)
        // Agregar filas del bono: iterar unidades con el mismo folder de bono
        const bonoFolders = nominaFolder.getFoldersByName(bonoNombre);
        if (bonoFolders.hasNext()) {
          const bonoFolder      = bonoFolders.next();
          const bonoUnidFolders = bonoFolder.getFolders();
          while (bonoUnidFolders.hasNext()) {
            const bonoUnidFolder = bonoUnidFolders.next();
            const unidadName2    = bonoUnidFolder.getName().trim();
            const folderId2      = bonoUnidFolder.getId();
            const creds2         = unitMap[unidadName2];
            if (!creds2) continue;

            rows.push([
              "",
              TRIMESTRE_2T,
              unidadName2,
              creds2.usuario,
              creds2.pass,
              qnaLabel,
              nominaLabel,
              bonoNombre,
              folderId2,
              "✓ Asignada"
            ]);
            contador++;
          }
        }
      }
    }
  }

  if (rows.length === 0) {
    Logger.log("❌ No se encontraron filas. Verifica la estructura de carpetas.");
    return;
  }

  // Escribir en el Sheet a partir de la última fila
  const startRow = lastRow + 1;
  sheet.getRange(startRow, 1, rows.length, 10).setValues(rows);

  // Limpiar caché para que el sistema actualice
  limpiarCache();

  Logger.log("✅ Filas agregadas: " + contador);
  Logger.log("✅ Desde fila: " + startRow + " hasta: " + (startRow + rows.length - 1));
}

// ── Normalizar nombre de QNA ─────────────────────────────────
function normalizarQNA(nombre) {
  // "QNA 07" → "QNA 07", "QNA07" → "QNA 07"
  const match = nombre.match(/(\d+)/);
  if (match) {
    const num = parseInt(match[1]);
    return "QNA " + (num < 10 ? "0" + num : num);
  }
  return nombre;
}

// ── Extraer nombre de nómina sin número prefijo ──────────────
function extraerNomina(nombre) {
  // "1. Ord" → "Ord", "2. Extras" → "Extras", "3. Bono Día Madres" → "Bono Día Madres"
  return nombre.replace(/^\d+\.\s*/, "").trim();
}
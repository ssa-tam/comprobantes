// ==============================================================
//  SST — Limpiar columna G (Nómina) del 3° Trimestre
//  Quita el prefijo "1. " de "1. Ord" → "Ord"
//  Corre limpiarNomina3T() una sola vez
// ============================================================

function limpiarNomina3T() {
  const sheet   = SpreadsheetApp.openById("1PyjT6E8FOiaeH79eKCMKIOiQbpmKqcpoyCfVr7Sd9ww").getSheetByName("Accesos");
  const lastRow = sheet.getLastRow();
  const data    = sheet.getRange(4, 1, lastRow - 3, 10).getValues();

  let corregidas = 0;
  const updates  = [];

  for (let i = 0; i < data.length; i++) {
    const trim   = (data[i][1] || "").toString().trim();
    const nomina = (data[i][6] || "").toString().trim();

    if (trim === "3° Trimestre") {
      // Quitar prefijo numérico: "1. Ord" → "Ord", "2. Extras" → "Extras"
      const limpia = nomina.replace(/^\d+\.\s*/, "").trim();
      if (limpia !== nomina) {
        updates.push({ fila: i + 4, valor: limpia });
        corregidas++;
      }
    }
  }

  Logger.log("Filas a corregir: " + corregidas);

  // Actualizar en lotes
  for (const u of updates) {
    sheet.getRange(u.fila, 7).setValue(u.valor);
  }

  // Limpiar caché
  const cache = CacheService.getScriptCache();
  cache.remove("sst_idx_v9_meta");
  cache.remove("sst_ctrl_v9");

  Logger.log("✅ Corrección completada: " + corregidas + " filas actualizadas");
}
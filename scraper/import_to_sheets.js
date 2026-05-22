/**
 * Google Apps Script – Importa products.xlsx a Google Sheets
 * ============================================================
 * USO:
 *  1. Sube products.xlsx a Google Drive
 *  2. Abre el archivo como Google Sheet (Archivo > Abrir con > Hojas de Cálculo)
 *  3. Herramientas > Editor de Apps Script
 *  4. Pega este código y ejecuta setupDataValidation()
 * ============================================================
 */

const SHEET_NAMES = {
  PRODUCTS:  "Productos",
  VARIANTS:  "Variantes",
  CATS:      "Categorias",
  IMAGES:    "Imagenes",
  MENUS:     "Menus",
  TEXTS:     "Textos_Banners",
  ASSETS:    "Assets_Tech",
};

/** Aplica formato profesional a todas las hojas */
function setupDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.values(SHEET_NAMES).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;

    // Cabecera
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground("#1F4E79").setFontColor("#FFFFFF")
               .setFontWeight("bold").setFontSize(11)
               .setHorizontalAlignment("center");

    // Filas alternas
    if (lastRow > 1) {
      for (let r = 2; r <= lastRow; r++) {
        const bg = (r % 2 === 0) ? "#EEF4FA" : "#FFFFFF";
        sheet.getRange(r, 1, 1, lastCol).setBackground(bg);
      }
    }

    // Freeze fila 1
    sheet.setFrozenRows(1);

    // Filtros
    sheet.getRange(1, 1, lastRow, lastCol).createFilter();

    // Auto-resize columnas
    for (let c = 1; c <= lastCol; c++) {
      sheet.autoResizeColumn(c);
      const w = sheet.getColumnWidth(c);
      if (w > 400) sheet.setColumnWidth(c, 400);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("✓ Formato aplicado correctamente", "Listo", 3);
}

/** Genera un menú personalizado al abrir el archivo */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🛒 ProductosLimeza")
    .addItem("Aplicar formato", "setupDataValidation")
    .addItem("Exportar productos a JSON", "exportProductsJSON")
    .addToUi();
}

/** Exporta la hoja Productos como JSON al log */
function exportProductsJSON() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  if (!sheet) { Logger.log("Hoja Productos no encontrada"); return; }

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const json = rows.map(row =>
    Object.fromEntries(headers.map((h, i) => [h, row[i]]))
  );

  Logger.log(JSON.stringify(json, null, 2));
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${rows.length} productos exportados al log`, "JSON Export", 5
  );
}

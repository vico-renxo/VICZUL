var SPREADSHEET_ID = '1c04OYCLKtWBV54WB--drcK58e629mfB_bmXwzb6wRlc';
var SHEET_NAME = 'DATA';

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('VICZUL - Productos de Limpieza')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function buildSlideTrack() {
  var html = '';
  var base = 'https://productosdelimpiezalima.com/image/logos_clientes/';
  for (var pass = 0; pass < 2; pass++) {
    for (var i = 1; i <= 45; i++) {
      html += '<div class="slide"><img src="' + base + i + '.webp" alt="" loading="lazy"></div>\n';
    }
  }
  return html;
}

function getProductData() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { products: [], categories: {} };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { products: [], categories: {} };

    var products   = [];
    var categories = {};
    var seen       = {};

    for (var i = 1; i < data.length; i++) {
      var row    = data[i];
      var nombre = String(row[3] || '').trim();
      if (!nombre) continue;

      var cat    = String(row[1]  || '').trim();
      var subcat = String(row[2]  || '').trim();
      var sinIgv = parseFloat(row[4]) || 0;
      var conIgv = parseFloat(row[5]) || 0;

      // Columnas opcionales (agregar a la hoja si se necesitan)
      var imagen        = String(row[6]  || '').trim();        // G: URL de imagen
      var precioAntes   = parseFloat(row[7])  || 0;           // H: Precio anterior (para descuento)
      var liquidacion   = String(row[8]  || '').toUpperCase() === 'SI'; // I: Liquidación
      var popular       = String(row[9]  || '').toUpperCase() === 'SI'; // J: Popular / Más vendido
      var sku           = String(row[10] || '').trim();        // K: SKU
      var sinStock      = String(row[11] || '').toUpperCase() === 'NO'; // L: Sin stock

      products.push({
        num:          Number(row[0]) || i,
        categoria:    cat,
        subcategoria: subcat,
        nombre:       nombre,
        precioSinIgv: sinIgv,
        precioConIgv: conIgv,
        imagen:       imagen,
        precioAntes:  precioAntes,
        liquidacion:  liquidacion,
        popular:      popular,
        sku:          sku,
        enStock:      !sinStock
      });

      if (cat) {
        if (!categories[cat]) categories[cat] = [];
        var key = cat + '||' + subcat;
        if (subcat && !seen[key]) {
          seen[key] = true;
          categories[cat].push(subcat);
        }
      }
    }

    for (var k in categories) categories[k].sort();

    return { products: products, categories: categories };

  } catch (err) {
    Logger.log('Error: ' + err);
    return { error: err.toString(), products: [], categories: {} };
  }
}

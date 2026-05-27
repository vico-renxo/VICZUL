var SPREADSHEET_ID = '1c04OYCLKtWBV54WB--drcK58e629mfB_bmXwzb6wRlc';
var SHEET_NAME     = 'DATA';
var CONFIG_NAME    = 'CONFIG';
var ICONS_NAME     = 'ICONOS_CAT';

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

/* ============================================================
   CONFIG — lee la pestaña CONFIG de la hoja de cálculo
   ============================================================ */
function getConfigFromSheet(ss) {
  var defaults = {
    WA_NUMBER:        '51999999999',
    EMPRESA:          'VICZUL',
    EMPRESA_DESC:     'Productos de Limpieza',
    SLOGAN:           'Hasta 30% de descuento',
    MES:              'Mayo 2026',
    BANNER_URL:       '',
    BANNER_VIDEO:     '',
    BANNER_1:         '',
    BANNER_2:         '',
    BANNER_3:         '',
    BANNER_4:         '',
    BANNER_5:         '',
    BANNER_INTERVALO: '4500',
    FACEBOOK:         '#',
    INSTAGRAM:        '#',
    TIKTOK:           '#',
    WHATSAPP_MSG:     'Hola, quisiera hacer un pedido.'
  };
  try {
    var cfgSheet = ss.getSheetByName(CONFIG_NAME);
    if (!cfgSheet) return defaults;
    var rows = cfgSheet.getDataRange().getValues();
    var config = {};
    for (var i = 0; i < rows.length; i++) {
      var key = String(rows[i][0] || '').trim().toUpperCase();
      var val = String(rows[i][1] || '').trim();
      if (!key || key.charAt(0) === '#' || key === 'CLAVE') continue;
      config[key] = val;
    }
    for (var k in defaults) {
      if (config[k] === undefined || config[k] === '') config[k] = defaults[k];
    }
    // Convertir URLs de banner a formato directo de imagen/video
    var bannerKeys = ['BANNER_URL','BANNER_VIDEO','BANNER_1','BANNER_2','BANNER_3','BANNER_4','BANNER_5'];
    for (var b = 0; b < bannerKeys.length; b++) {
      if (config[bannerKeys[b]]) config[bannerKeys[b]] = toDriveImgUrl(config[bannerKeys[b]]);
    }
    return config;
  } catch (e) {
    Logger.log('getConfigFromSheet error: ' + e);
    return defaults;
  }
}

/* ============================================================
   setupConfigSheet — crea/reinicia la pestaña CONFIG
   Ejecutar una sola vez desde el editor de GAS: menu Ejecutar → setupConfigSheet
   ============================================================ */
function setupConfigSheet() {
  try {
    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var cfgSheet = ss.getSheetByName(CONFIG_NAME);
    if (!cfgSheet) {
      cfgSheet = ss.insertSheet(CONFIG_NAME);
    } else {
      cfgSheet.clearContents();
      cfgSheet.clearFormats();
    }

    var rows = [
      ['CLAVE',        'VALOR',                    'DESCRIPCIÓN (no editar esta columna)'],
      ['WA_NUMBER',    '51999999999',               'Número WhatsApp con código de país (sin +). Ej: 51987654321'],
      ['EMPRESA',      'VICZUL',                   'Nombre de tu empresa (aparece en el logo y el footer)'],
      ['EMPRESA_DESC', 'Productos de Limpieza',    'Descripción corta bajo el nombre'],
      ['SLOGAN',       'Hasta 30% de descuento',   'Texto de oferta en la barra superior'],
      ['MES',          'Mayo 2026',                'Mes/período de la promoción. Ej: Junio 2026'],
      ['BANNER_VIDEO',     '',     'URL o ID Drive del VIDEO principal del banner (MP4). Tiene prioridad sobre imágenes.'],
      ['BANNER_1',        '',     'Imagen 1 del slideshow — URL o ID de Google Drive'],
      ['BANNER_2',        '',     'Imagen 2 del slideshow — URL o ID de Google Drive'],
      ['BANNER_3',        '',     'Imagen 3 del slideshow — URL o ID de Google Drive'],
      ['BANNER_4',        '',     'Imagen 4 del slideshow (opcional)'],
      ['BANNER_5',        '',     'Imagen 5 del slideshow (opcional)'],
      ['BANNER_INTERVALO','4500', 'Tiempo entre imágenes en milisegundos (4500 = 4.5 seg)'],
      ['FACEBOOK',     '#',                        'URL completa de tu página de Facebook'],
      ['INSTAGRAM',    '#',                        'URL completa de tu Instagram'],
      ['TIKTOK',       '#',                        'URL completa de tu TikTok'],
      ['WHATSAPP_MSG', 'Hola, quisiera hacer un pedido.', 'Mensaje inicial de WhatsApp (sin pedido en el carrito)'],
    ];

    cfgSheet.getRange(1, 1, rows.length, 3).setValues(rows);

    // Formato encabezado
    var hdrRange = cfgSheet.getRange(1, 1, 1, 3);
    hdrRange.setBackground('#004EA5').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);

    // Formato columna CLAVE (bold, fondo claro)
    cfgSheet.getRange(2, 1, rows.length - 1, 1)
      .setBackground('#EEF3FF').setFontWeight('bold').setFontSize(10);

    // Ancho de columnas
    cfgSheet.setColumnWidth(1, 160);
    cfgSheet.setColumnWidth(2, 320);
    cfgSheet.setColumnWidth(3, 420);

    // Proteger columna de descripción (solo advertencia, no bloqueo)
    cfgSheet.getRange(1, 3, rows.length, 1).setFontColor('#888888').setFontStyle('italic');

    Logger.log('✅ Pestaña CONFIG creada correctamente en la hoja de cálculo.');
    return '✅ CONFIG creada. Ahora edita los valores en la columna VALOR de la pestaña CONFIG.';
  } catch (e) {
    Logger.log('❌ Error en setupConfigSheet: ' + e);
    return '❌ Error: ' + e.toString();
  }
}

/* ============================================================
   setupIconsCatSheet — crea/reinicia la pestaña ICONOS_CAT
   Ejecutar una sola vez desde el editor GAS: Ejecutar → setupIconsCatSheet
   ============================================================ */
function setupIconsCatSheet() {
  try {
    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = ss.getSheetByName(SHEET_NAME);
    if (!dataSheet) return '❌ No se encontró la pestaña DATA.';

    // Leer categorías únicas de DATA (col B = índice 1)
    var data = dataSheet.getDataRange().getValues();
    var catsSet = {};
    for (var i = 1; i < data.length; i++) {
      var cat = String(data[i][1] || '').trim();
      if (cat) catsSet[cat] = true;
    }
    var cats = Object.keys(catsSet).sort();

    // Crear o limpiar pestaña ICONOS_CAT
    var iconSheet = ss.getSheetByName(ICONS_NAME);
    if (!iconSheet) {
      iconSheet = ss.insertSheet(ICONS_NAME);
    } else {
      iconSheet.clearContents();
      iconSheet.clearFormats();
    }

    // Encabezado + filas
    var rows = [['CATEGORÍA', 'URL ICONO (imagen)', 'NOTAS']];
    cats.forEach(function(cat) {
      rows.push([cat, '', 'Pegar aquí la URL de la imagen (PNG/WebP/SVG)']);
    });

    iconSheet.getRange(1, 1, rows.length, 3).setValues(rows);

    // Formato encabezado
    iconSheet.getRange(1, 1, 1, 3)
      .setBackground('#004EA5').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(11);

    // Formato categorías
    iconSheet.getRange(2, 1, cats.length, 1)
      .setBackground('#EEF3FF').setFontWeight('bold').setFontSize(10);

    // Formato notas
    iconSheet.getRange(2, 3, cats.length, 1)
      .setFontColor('#888888').setFontStyle('italic');

    // Ancho columnas
    iconSheet.setColumnWidth(1, 220);
    iconSheet.setColumnWidth(2, 380);
    iconSheet.setColumnWidth(3, 300);

    Logger.log('✅ Pestaña ICONOS_CAT creada con ' + cats.length + ' categorías.');
    return '✅ Pestaña ICONOS_CAT creada con ' + cats.length + ' categorías. Pega las URLs en la columna B.';
  } catch (e) {
    Logger.log('❌ Error en setupIconsCatSheet: ' + e);
    return '❌ Error: ' + e.toString();
  }
}

/* ============================================================
   toDriveImgUrl — acepta ID puro, link de Drive o URL normal
   y siempre devuelve una URL de imagen directa
   ============================================================ */
function toDriveImgUrl(val) {
  if (!val) return '';
  val = val.trim();
  // URL directa que no es de Drive: se usa tal cual
  if (val.indexOf('http') === 0 && val.indexOf('drive.google.com') === -1) return val;
  // Extraer ID de cualquier link de Drive (formato /d/ID/)
  var m = val.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return 'https://lh3.googleusercontent.com/d/' + m[1];
  // ID puro: solo caracteres alfanuméricos + _ - y longitud típica de Drive (≥20)
  if (/^[a-zA-Z0-9_-]{20,}$/.test(val)) return 'https://lh3.googleusercontent.com/d/' + val;
  return val;
}

/* ============================================================
   getProductData — productos + categorías + config
   ============================================================ */
function getProductData() {
  try {
    var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    var config = getConfigFromSheet(ss);
    var sheet  = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { products: [], categories: {}, config: config };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { products: [], categories: {}, config: config };

    var products   = [];
    var categories = {};
    var seen       = {};

    for (var i = 1; i < data.length; i++) {
      var row    = data[i];
      var nombre = String(row[3] || '').trim();
      if (!nombre) continue;

      var cat    = String(row[1]  || '').trim();
      var subcat = String(row[2]  || '').trim();
      var sinIgv = parseFloat(row[4])  || 0;
      var conIgv = parseFloat(row[5])  || 0;

      var imagen      = String(row[6]  || '').trim();
      var precioAntes = parseFloat(row[7])  || 0;
      var liquidacion = String(row[8]  || '').toUpperCase() === 'SI';
      var popular     = String(row[9]  || '').toUpperCase() === 'SI';
      var sku         = String(row[10] || '').trim();
      var sinStock    = String(row[11] || '').toUpperCase() === 'NO';

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
        if (subcat && !seen[key]) { seen[key] = true; categories[cat].push(subcat); }
      }
    }

    for (var k in categories) categories[k].sort();

    // Leer iconos de categoría desde ICONOS_CAT
    var catIcons = {};
    try {
      var iconSheet = ss.getSheetByName(ICONS_NAME);
      if (iconSheet) {
        var iconData = iconSheet.getDataRange().getValues();
        for (var j = 1; j < iconData.length; j++) {
          var catName = String(iconData[j][0] || '').trim();
          var iconUrl = toDriveImgUrl(String(iconData[j][1] || '').trim());
          if (catName && iconUrl) catIcons[catName] = iconUrl;
        }
      }
    } catch (iconErr) {
      Logger.log('catIcons read error: ' + iconErr);
    }

    return { products: products, categories: categories, config: config, catIcons: catIcons };

  } catch (err) {
    Logger.log('getProductData error: ' + err);
    return { error: err.toString(), products: [], categories: {}, config: {} };
  }
}

# Extractor Completo – productosdelimpiezalima.com

Sistema de extracción de datos completo para el sitio WooCommerce.

## Estructura

```
VICZUL/
├── scraper/
│   ├── woo_extractor.py        ← Script principal de extracción
│   ├── config.py               ← Configuración (API keys aquí)
│   ├── import_to_sheets.js     ← Google Apps Script para Sheets
│   └── requirements.txt        ← Dependencias Python
├── data/                       ← Archivos generados (tras ejecutar)
│   ├── products_<ts>.xlsx      ← Excel con 7 hojas
│   ├── database_<ts>.sql       ← SQL listo para importar
│   ├── products.json
│   ├── categories.json
│   ├── images.json
│   └── sitemap.json
└── assets/                     ← Imágenes descargadas (opcional)
```

## Configuración rápida

### 1. Obtener credenciales WooCommerce API

1. Entra a tu admin de WordPress
2. Ve a **WooCommerce > Ajustes > Avanzado > API REST**
3. Clic en **Agregar clave**
4. Descripción: `Extractor`, Usuario: tu admin, Permisos: **Lectura**
5. Copia `consumer_key` y `consumer_secret`

### 2. Configurar el script

Edita `scraper/config.py`:

```python
WC_CONSUMER_KEY    = "ck_TU_CLAVE_AQUI"
WC_CONSUMER_SECRET = "cs_TU_SECRET_AQUI"
```

### 3. Ejecutar

```bash
cd scraper
pip install -r requirements.txt
python woo_extractor.py
```

Los archivos se generan en `/data/`.

## Excel generado (`products_<timestamp>.xlsx`)

| Hoja | Contenido |
|------|-----------|
| **Productos** | nombre, precio, SKU, stock, categorías, descripción, URL imagen |
| **Variantes** | variantes de productos variables con precios y stock |
| **Categorias** | id, nombre, padre, URL, conteo |
| **Imagenes** | producto_id, URL, alt text, posición |
| **Menus** | texto, URL, nivel de profundidad |
| **Textos_Banners** | páginas, banners, contenido HTML parseado |
| **Assets_Tech** | tecnología, colores HEX, fuentes |

## SQL generado (`database_<timestamp>.sql`)

Tablas creadas:
- `categories`
- `products`
- `product_images`
- `product_variations`
- `nav_menus`
- `pages`

## Google Sheets

1. Sube `products_<ts>.xlsx` a Google Drive
2. Ábrelo con Hojas de Cálculo de Google
3. Ve a **Herramientas > Editor de Apps Script**
4. Pega el contenido de `scraper/import_to_sheets.js`
5. Ejecuta `setupDataValidation()` para aplicar formato

## Sitio mapeado

- **URL**: https://productosdelimpiezalima.com
- **Plataforma**: WordPress + WooCommerce
- **API**: WooCommerce REST API v3

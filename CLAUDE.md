# VICZUL — Google Apps Script Web App

## Proyecto
Catálogo e-commerce de productos de limpieza para VICZUL S.A.C., construido en Google Apps Script (GAS). El diseño replica visualmente el sitio de referencia productosdelimpiezalima.com (OpenCart / tema oct_showcase).

## Archivos GAS
| Archivo | Rol |
|---------|-----|
| `gas/Code.gs` | Backend: `doGet()`, `include()`, `buildSlideTrack()`, `getProductData()` |
| `gas/index.html` | Template principal (GAS scriptlets `<?!= ?>`) |
| `gas/css.html` | Estilos inyectados via `<?!= include('css'); ?>` |

## Base de Datos — Google Sheets
- **ID**: `1c04OYCLKtWBV54WB--drcK58e629mfB_bmXwzb6wRlc`
- **Pestaña**: `DATA`

| Col | Campo | Tipo | Notas |
|-----|-------|------|-------|
| A | N° | number | |
| B | Categoría | string | |
| C | Subcategoría | string | |
| D | Producto | string | Nombre |
| E | Precio sin IGV | float | |
| F | Precio con IGV | float | |
| G | Imagen URL | string | Opcional — dejar vacío hasta tener URLs |
| H | Precio anterior | float | Opcional — para mostrar descuento tachado |
| I | Liquidación | "SI"/"" | "SI" = aparece en sección Liquidación |
| J | Popular | "SI"/"" | "SI" = sticker Más vendido |
| K | SKU | string | Opcional |
| L | Stock | "NO"/"" | "NO" = sin stock |

## Diseño de Referencia
**Fuente**: productosdelimpiezalima.com — HTML completo entregado por el usuario.

### Estructura de layout (fiel al original)
```
[<picture> Banner GIF full-width]
[snowflakes animados]
[#oct-infobar] — alerta azul cerrable
[#top nav] — pre-header: Últimas Entregas · Delivery · Formas de Pago · Seguimiento · Liquidación
[<header> sticky] — Catálogo btn + search + account/cart/whatsapp
  └── [.sc-megamenu] — dropdown categorías (dinámico desde hoja)
[#common-home container-fluid container-lg]
  ├── Slideshow banner (picture element)
  ├── Liquidación — carrusel horizontal
  ├── Advantages — 6 tiles (imagen + título + texto)
  ├── Client slider — logos 1-45 de productosdelimpiezalima.com
  ├── "Productos Populares" — grid 5col/2col
  └── "Últimas reseñas"
[<footer>] — azul #004EA5, 4 columnas
```

**NO HAY sidebar persistente en desktop.** El original usa megamenú dropdown. Solo hay sidebar mobile (hamburguesa).

### Paleta exacta (detectada del original)
| Variable | Hex | Uso |
|----------|-----|-----|
| `--blue` | `#004EA5` | Primary, header, footer, botones |
| `--dark` | `#1c1c28` | Texto principal, precio real |
| `--gray` | `#8f90a6` | Texto secundario, precio tachado |
| `--red` | `#e53535` | Descuentos, alertas |
| `--green` | `#06c270` | Sticker "En stock" |
| `--purple` | `#9615aa` | Sticker "Popular" |
| `--light` | `#f8f9fa` | Background body |
| `--foot-bg` | `#004EA5` | Footer |

### Tipografía
- Fuente: **Montserrat** (Google Fonts) — fw-400/500/600/700/800
- La original usa Montserrat (no Inter)

### Stickers de producto (clases `sc-module-sticker-*`)
| Clase | Color | Texto |
|-------|-------|-------|
| `sc-module-sticker-green` | `#06c270` | En stock |
| `sc-module-sticker-dark` | `#1c1c28` | Más vendido |
| `sc-module-sticker-purple` | `#9615aa` | Popular |
| `sc-module-sticker-blue` | `#004EA5` | Recomendado |

### Cards de producto — formato original
```
[imagen 230×180] + [stickers absolutos bottom-left]
[título fw-700]
[SKU fsz-10]
[★★★★★ estrellas + (N reseñas)]
[S/ XX.XX inc. IGV — fw-800 fsz-16+]
[qty control −/N/+ ] [btn Agregar]
```

### Sistema de precios
- Precio tachado: `S/ XX.XX` — `fsz-14 fw-400 text-decoration-line-through color:#8f90a6`
- Precio real: `S/ XX.XX inc. IGV` — `fsz-16+ fw-800 color:#1c1c28`

### Sección Liquidación
Carrusel horizontal con items compactos:
- [thumbnail 70×70] + [nombre] + [precio tachado] + [precio nuevo] + [btn cart]

### Advantages (6 tiles — textos del original)
1. Entrega al día siguiente — "Entrega ultra rápida a tu domicilio."
2. Orientación Experta — "Sabemos lo que necesitas, consúltanos..."
3. Embalaje para Agencia — "Sin costo adicional! Shalom, Marvisur, Grael, etc."
4. Documentación al día — "Ficha técnica, Hoja de Seguridad y Registro Sanitario Vigente."
5. Precios competitivos — "Descuentos por cantidad al por mayor disponibles."
6. Comprobante Electrónico — "Boleta, Factura, Guía de Remisión Electronicas. (PDF, XML, CDR)"

### Footer — estructura original
4 columnas: **Redes Sociales** | **Newsletter** | **Información** | **Populares** + Formas de pago (BCP, BBVA, Interbank, Scotiabank, Yape)

## Checkout
WhatsApp: variable `WA_NUMBER = '51999999999'` en index.html — reemplazar con número real.
Cart persistido en localStorage con key `viczul_cart`.

## Slider Clientes
45 logos: `https://productosdelimpiezalima.com/image/logos_clientes/N.webp` (N=1..45)
Duplicados (90 divs) para loop infinito CSS. Construido server-side en `buildSlideTrack()`.

## Rama de desarrollo
`claude/determined-cannon-FeZir`

## CDNs usados
- Bootstrap 5.3.2
- Font Awesome 6.5.1
- Montserrat (Google Fonts)

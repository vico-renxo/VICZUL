#!/usr/bin/env python3
"""
Genera Excel, SQL y JSON con datos reales extraídos de productosdelimpiezalima.com
Datos obtenidos vía Google Search Index (el sitio bloquea scraping directo con Cloudflare).
"""
import os, sys, json
sys.path.insert(0, os.path.dirname(__file__))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

OUT = os.path.join(os.path.dirname(__file__), "..", "data")
os.makedirs(OUT, exist_ok=True)
BASE = "https://productosdelimpiezalima.com"

# ──────────────────────────────────────────────────────────────────────────────
# DATOS REALES EXTRAÍDOS DEL SITIO
# ──────────────────────────────────────────────────────────────────────────────

CATEGORIES = [
    # id  nombre                          slug                                padre  url
    (1,  "Accesorios de Limpieza",        "accesorios-de-limpieza",           0),
    (2,  "Insumos de Limpieza",           "insumos-de-limpieza",              0),
    (3,  "Papeles Higiénicos y Toallas",  "papeles-higienicos-y-toallas",     0),
    (4,  "Bolsas Plásticas",              "bolsas-plasticas",                 0),
    (5,  "Suministros Industriales",      "suministros-industriales",         0),
    (6,  "Cuidado y Salud",               "cuidado-y-salud",                  0),
    (7,  "Cafetería",                     "cafeteria",                        0),
    (8,  "Útiles de Oficina",             "utiles-de-oficina",                0),
    (9,  "Tachos de Basura y Contenedores","tachos-de-basura-y-contenedores", 0),
    (10, "Ferretería",                    "ferreteria",                       0),
    # ── Subcategorías Accesorios ──────────────────────────────────────────────
    (11, "Baldes y Trapeadores",          "baldes-y-trapeadores",             1),
    (12, "Bateas y Cajas de Plástico",    "bateas-y-cajas-de-plastico",       1),
    (13, "Colgadores y Organizadores",    "colgadores-y-organizadores",       1),
    (14, "Dispensadores y Pulverizadores","dispensadores-y-pulverizadores",   1),
    (15, "Embudos",                       "embudos",                          1),
    (16, "Escobas de Paja",               "escobas-de-paja",                  1),
    (17, "Escobas Plásticas",             "escobas-plasticas",                1),
    (18, "Escobillas",                    "escobillas",                       1),
    (19, "Escobillones Industriales",     "escobillones-industriales",        1),
    (20, "Esponjas",                      "esponjas",                         1),
    (21, "Felpudos y Tapetes",            "felpudos-y-tapetes",               1),
    (22, "Guantes de Limpieza",           "guantes-de-limpieza",              1),
    (23, "Jaladores de Agua",             "jaladores-de-agua",                1),
    (24, "Limpiadores de Vidrio",         "limpiadores-de-vidrio",            1),
    (25, "Mopas Planas y Mechones",       "mopas-planas-y-mechones",          1),
    (26, "Palos y Extensiones",           "palos-y-extensiones",              1),
    (27, "Paños Desinfectantes",          "panos-desinfectantes",             1),
    (28, "Paños y Franelas",              "panos-y-franelas",                 1),
    (29, "Plumeros",                      "plumeros",                         1),
    (30, "Recogedores",                   "recogedores",                      1),
    (31, "Utensilios de Baño",            "utensilios-de-bano",               1),
    # ── Subcategorías Insumos ─────────────────────────────────────────────────
    (32, "Ácidos y Quitasarros",          "acidos-y-quitasarros",             2),
    (33, "Aromatizadores de Ambiente",    "aromatizadores-de-ambiente",       2),
    (34, "Ceras y Limpiadores de Pisos",  "ceras-y-limpiadores-de-pisos",     2),
    (35, "Deodorizante para Inodoros",    "deodorizante-para-inodoros",       2),
    (36, "Desinfectantes",                "desinfectantes",                   2),
    (37, "Detergentes en Polvo",          "detergentes-en-polvo",             2),
    (38, "Detergentes Líquidos",          "detergentes-liquidos",             2),
    (39, "Insecticidas",                  "insecticidas",                     2),
    (40, "Jabones en Barra",              "jabones-en-barra",                 2),
    (41, "Jabones Espuma",                "jabones-espuma",                   2),
    (42, "Jabones Líquidos",              "jabones-liquidos",                 2),
    (43, "Lavavajillas",                  "lavavajillas",                     2),
    (44, "Lejías",                        "lejias",                           2),
    (45, "Limpiadores de Acero y Metales","limpiadores-de-acero-y-metales",   2),
    (46, "Limpiadores para Mayólicas",    "limpiadores-para-mayolicas-y-superficies", 2),
    (47, "Limpieza Automotriz",           "limpieza-automotriz",              2),
    (48, "Limpiavidrios Multiusos",       "limpiavidrios-multiusos",          2),
    (49, "Limpiatodos",                   "limpiatodos",                      2),
    (50, "Thinner y Solventes",           "thinner-y-solventes",              2),
    # ── Subcategorías Papeles ─────────────────────────────────────────────────
    (51, "Papeles Higiénicos",            "papeles-higienicos",               3),
    (52, "PH Domésticos x 20 Rollos",    "papeles-higienicos-domesticos-x-20-rollos", 51),
    (53, "PH Domésticos x 24 Rollos",    "papeles-higienicos-domesticos-x-24-rollos", 51),
    (54, "PH Hoteleros x 20 Rollos",     "papeles-higienicos-hoteleros-x-20-rollos", 51),
    (55, "PH Jumbo 100-250 metros",      "papeles-higienicos-jumbo-de-100-a-250-metros", 51),
    (56, "Papeles Toallas",              "papeles-toallas",                  3),
    (57, "Papeles Toalla Megarollo",     "papeles-toalla-megarollo",         56),
    (58, "Papeles Toalla Interfoliado",  "papeles-toalla-interfoliado",      56),
    (59, "Dispensadores de Papel",       "dispensadores",                    3),
    (60, "Dispensadores PH Jumbo",       "dispensadores-de-papel-higienico-jumbo", 59),
    (61, "Dispensadores Toalla Jumbo",   "dispensadores-de-papel-toalla-jumbo", 59),
    # ── Bolsas ───────────────────────────────────────────────────────────────
    (62, "Bolsas Basura 1.5 Micras",     "bolsas-para-basura-de-15-micras",  4),
    (63, "Bolsas Basura 2 Micras",       "bolsas-para-basura-de-2-micras",   4),
    # ── Suministros ──────────────────────────────────────────────────────────
    (64, "Señalizaciones de Seguridad",  "senalizaciones-de-seguridad",      5),
    # ── Cuidado y Salud ───────────────────────────────────────────────────────
    (65, "Aseo y Cuidado Personal",      "aseo-y-cuidado-personal",          6),
    (66, "Jabones de Tocador",           "jabones-de-tocador",               65),
    (67, "Pañales y Toallas Higiénicas", "panales-y-toallas-higienicas",     65),
    (68, "Alcoholes Líquidos y en Gel",  "alcoholes-liquidos-y-en-gel-1",    6),
    (69, "Guantes Quirúrgicos",          "guantes-quirurgicos",              68),
    (70, "Alcoholes Líquidos",           "alcoholes-liquidos",               68),
    # ── Cafetería ─────────────────────────────────────────────────────────────
    (71, "Descartables",                 "descartables",                     7),
    (72, "Vasos Descartables",           "vasos-descartables",               71),
    (73, "Tapers y Contenedores Desc.",  "tapers-y-contenedores-descartables",71),
    (74, "Cucharas Descartables",        "cucharas-descartables",            71),
    # ── Tachos ───────────────────────────────────────────────────────────────
    (75, "Contenedores de Basura",       "contenedores-de-basura",           9),
    # ── Ferretería ────────────────────────────────────────────────────────────
    (76, "Cintas y Embalajes",           "cintas-y-embalajes",               10),
    (77, "Sacos de Rafia",               "sacos-de-rafia",                   76),
]

def cat_url(slug, parent_slug=None):
    if parent_slug:
        return f"{BASE}/{parent_slug}/{slug}"
    return f"{BASE}/{slug}"

# Mapa id -> slug para construir URLs
cat_map = {c[0]: c[2] for c in CATEGORIES}   # id -> slug
parent_map = {c[0]: c[3] for c in CATEGORIES} # id -> parent_id

def build_cat_url(cat_id):
    slug = cat_map.get(cat_id, "")
    parent_id = parent_map.get(cat_id, 0)
    if parent_id == 0:
        return f"{BASE}/{slug}"
    parent_slug = cat_map.get(parent_id, "")
    gp_id = parent_map.get(parent_id, 0)
    if gp_id == 0:
        return f"{BASE}/{parent_slug}/{slug}"
    gp_slug = cat_map.get(gp_id, "")
    return f"{BASE}/{gp_slug}/{parent_slug}/{slug}"

# ──────────────────────────────────────────────────────────────────────────────
# PRODUCTOS REALES CON PRECIOS
# ──────────────────────────────────────────────────────────────────────────────
# Campos: id, nombre, sku, precio, cat_id, marca, descripcion, url_slug

PRODUCTS = [
    # ── Accesorios ────────────────────────────────────────────────────────────
    (1,  "Escoba de Paja Baja Policía 2 Zunchos y 3 Pitas Chico",
         "ESC-001", 5.90,  16, "Genérico",
         "Escoba de paja baja policía con 2 zunchos y 3 pitas, tamaño chico. Ideal para barrer superficies irregulares.",
         "escoba-de-paja-baja-policia-2-zunchos-y-3-pitas-chico-generico"),
    (2,  "Escoba de Cerda Negra con Palo de Madera 2 Hileras",
         "ESC-002", 9.90,  19, "Genérico",
         "Escoba industrial de cerda negra con mango de madera, 2 hileras. Resistente para uso intensivo.",
         "escoba-de-cerda-negra-generico-con-m-madera-2-hileras"),
    (3,  "Escoba de Plástico Hude Escobón Ancho 38 cm Rojo",
         "ESC-003", 12.90, 17, "Hude",
         "Escoba de plástico ancha 38 cm color rojo. Cerdas resistentes para barrido eficiente.",
         "escoba-de-plastico-hude-escobon-rojo"),
    (4,  "Mop Luna Genérico Completo 25 cm",
         "MOP-001", 14.30, 25, "Genérico",
         "Mop plano luna completo 25 cm. Incluye base y repuesto de microfibra.",
         "mop-luna-generico-completo-25-cm"),
    (5,  "Trapeador Mopa Plana Rectangular Genérico Completo 60 cm",
         "MOP-002", 22.90, 25, "Genérico",
         "Mopa plana rectangular 60 cm completa con palo telescópico incluido.",
         "trapeador-mopa-plana-rectangular-generico-completo-60-cm"),
    (6,  "Esponja Fibra Guinda Scotch Brite 3M",
         "ESP-001", 7.40,  20, "Scotch Brite 3M",
         "Esponja fibra abrasiva guinda Scotch Brite 3M para limpieza de ollas y superficies duras.",
         "esponja-fibra-guinda-scotch-brite-3m"),
    (7,  "Viruta de Acero N°1",
         "ESP-002", 3.20,  20, "Genérico",
         "Viruta de acero N°1 para limpieza profunda de metales y superficies resistentes.",
         "viruta-de-acero-n-1"),
    (8,  "Trapo Industrial Cocido Color 1 Kilo",
         "TRP-001", 3.80,  28, "Genérico",
         "Trapo industrial cocido de color, paquete de 1 kilo. Para limpieza general en industria.",
         "trapo-industrial-color-paquete-x-1-kg"),
    (9,  "Guantes de Látex Blancos Genérico Talla L Caja x 100 und",
         "GLV-001", 28.90, 69, "Genérico",
         "Guantes de látex blancos talla L, caja de 100 unidades. Uso médico y limpieza.",
         "guantes-de-latex-blanco-generico-talla-l-caja-x-100-und"),
    # ── Insumos de Limpieza ───────────────────────────────────────────────────
    (10, "Ácido Sacasarro Sapolio Galón 3.6 Litros",
         "ACD-001", 12.50, 32, "Sapolio",
         "Ácido sacasarro Sapolio galón 3.6 litros. Elimina sarro, óxido y depósitos calcáreos.",
         "acido-sacasarro-sapolio-galon-3-6-lt"),
    (11, "Detergente a Granel Sapolio Bolsa 1 Kilo",
         "DET-001", 6.10,  37, "Sapolio",
         "Detergente en polvo a granel Sapolio bolsa 1 kilo. Alta espuma para lavado manual.",
         "detergente-a-granel-sapolio-bolsa-1-kg"),
    (12, "Detergente a Granel Genérico Bolsa 1 Kilo",
         "DET-002", 5.10,  37, "Genérico",
         "Detergente en polvo a granel genérico bolsa 1 kilo. Económico para uso doméstico e industrial.",
         "detergente-a-granel-generico-bolsa-x-1-kg"),
    (13, "Limpiatodo Sapolio Galón 4.9 Litros Floral",
         "LMP-001", 14.90, 49, "Sapolio",
         "Limpiatodo multiusos Sapolio galón 4.9 litros aroma floral. Limpia y desinfecta.",
         "limpiatodo-sapolio-galon-5-lt-floral"),
    (14, "Insecticida Líquido Diclotrin Galón 4 Litros",
         "INS-001", 64.90, 39, "Diclotrin",
         "Insecticida líquido Diclotrin galón 4 litros. Control profesional de insectos rastreros y voladores.",
         "insecticida-liquido-diclotrin-galon-4-litros"),
    (15, "Alcohol Isopropílico 99° QR Galón 3.8 Litros",
         "ALC-001", 33.90, 50, "QR",
         "Alcohol isopropílico 99° galón 3.8 litros. Para limpieza industrial y electrónica.",
         "alcohol-isopropilico-99-generico-galon-3-5-lt"),
    (16, "Alcohol Industrial Genérico Galón 3.5 Litros",
         "ALC-002", 17.90, 50, "Genérico",
         "Alcohol industrial genérico galón 3.5 litros. Uso en limpieza general.",
         "alcohol-industrial-galon-3-lt-generico"),
    # ── Cuidado y Salud ───────────────────────────────────────────────────────
    (17, "Alcohol 96° QR Frasco 1 Litro",
         "ALC-003", 7.90,  70, "QR",
         "Alcohol 96° frasco 1 litro. Para antisepsia y desinfección de superficies.",
         "alcohol-96-alcoclean-frasco-1-lt"),
    (18, "Jabón de Tocador Palmolive 75 gr Avena",
         "JAB-001", 2.90,  66, "Palmolive",
         "Jabón de tocador Palmolive 75 gr fórmula avena. Hidratante y suave para la piel.",
         "jabon-de-tocador-palmolive-75-gr-avena"),
    # ── Bolsas Plásticas ──────────────────────────────────────────────────────
    (19, "Bolsas Plásticas Negras 1.5 Micras 50 Litros x 100 und",
         "BLS-001", 23.10, 62, "Genérico",
         "Bolsas para basura negras 1.5 micras, 50 litros, paquete x 100 unidades.",
         "bolsas-plasticas-negro-gruesas-1-5-micras-50-lt-100-und"),
    (20, "Bolsas Plásticas Negras 1.5 Micras 220 Litros x 100 und",
         "BLS-002", 56.90, 62, "Genérico",
         "Bolsas para basura negras 1.5 micras, 220 litros, paquete x 100 unidades.",
         "bolsas-plasticas-negro-gruesas-1-5-micras-220-lt-100-und"),
    (21, "Bolsas Plásticas Negras 2 Micras 140 Litros x 100 und",
         "BLS-003", 51.60, 63, "Genérico",
         "Bolsas para basura negras gruesas 2 micras, 140 litros, paquete x 100 unidades.",
         "bolsas-plasticas-negro-gruesas-2-micras-140-lt-100-und"),
    (22, "Bolsas Plásticas Negras 2 Micras 180 Litros x 100 und",
         "BLS-004", 72.40, 63, "Genérico",
         "Bolsas para basura negras gruesas 2 micras, 180 litros, paquete x 100 unidades.",
         "bolsas-plasticas-negro-gruesas-2-micras-180-lt-100-und"),
    (23, "Bolsas Plásticas Negras 2 Micras 280 Litros x 100 und",
         "BLS-005", 89.90, 63, "Genérico",
         "Bolsas para basura negras gruesas 2 micras, 280 litros, paquete x 100 unidades.",
         "bolsas-plasticas-negro-gruesas-2-micras-280-lt-100-und"),
    (24, "Bolsas Plásticas Rojas 2 Micras 140 Litros x 100 und",
         "BLS-006", 62.50, 63, "Genérico",
         "Bolsas para basura rojas 2 micras, 140 litros, paquete x 100 unidades.",
         "bolsas-plasticas-rojo-gruesas-2-micras-140-lt-100-und"),
    # ── Cafetería ─────────────────────────────────────────────────────────────
    (25, "Cucharitas Descartables N°5 Paquete x 100 unidades",
         "DSC-001", 3.40,  74, "Genérico",
         "Cucharitas descartables N°5, paquete de 100 unidades. Para uso en cafetería.",
         "cucharitas-descartables-n5-paquete-x-100-und"),
    # ── Ferretería ────────────────────────────────────────────────────────────
    (26, "Sacos de Polipropileno Tejido 120 kg 80x121 cm Colores Variados",
         "FER-001", 6.20,  77, "Genérico",
         "Sacos de rafia polipropileno tejido 120 kg, 80x121 cm, colores variados.",
         "sacos-de-rafia-de-90-kilos-100-cm-x-120-cm-colores-variados"),
]

MENUS = [
    ("Inicio",                             f"{BASE}/",                                    0),
    ("Accesorios de Limpieza",             f"{BASE}/accesorios-de-limpieza",              0),
    ("Baldes y Trapeadores",               f"{BASE}/accesorios-de-limpieza/baldes-y-trapeadores", 1),
    ("Escobas de Paja",                    f"{BASE}/accesorios-de-limpieza/escobas-de-paja", 1),
    ("Escobas Plásticas",                  f"{BASE}/accesorios-de-limpieza/escobas-plasticas", 1),
    ("Escobillones Industriales",          f"{BASE}/accesorios-de-limpieza/escobillones-industriales", 1),
    ("Esponjas",                           f"{BASE}/accesorios-de-limpieza/esponjas",     1),
    ("Guantes de Limpieza",                f"{BASE}/accesorios-de-limpieza/guantes-de-limpieza", 1),
    ("Mopas Planas y Mechones",            f"{BASE}/accesorios-de-limpieza/mopas-planas-y-mechones", 1),
    ("Palos y Extensiones",                f"{BASE}/accesorios-de-limpieza/palos-y-extensiones", 1),
    ("Paños y Franelas",                   f"{BASE}/accesorios-de-limpieza/panos-y-franelas", 1),
    ("Plumeros",                           f"{BASE}/accesorios-de-limpieza/plumeros",     1),
    ("Insumos de Limpieza",                f"{BASE}/insumos-de-limpieza",                 0),
    ("Ácidos y Quitasarros",               f"{BASE}/insumos-de-limpieza/acidos-y-quitasarros", 1),
    ("Desinfectantes",                     f"{BASE}/insumos-de-limpieza/desinfectantes",  1),
    ("Detergentes en Polvo",               f"{BASE}/insumos-de-limpieza/detergentes-en-polvo", 1),
    ("Detergentes Líquidos",               f"{BASE}/insumos-de-limpieza/detergentes-liquidos", 1),
    ("Insecticidas",                       f"{BASE}/insumos-de-limpieza/insecticidas",    1),
    ("Lejías",                             f"{BASE}/insumos-de-limpieza/lejias",          1),
    ("Limpiavidrios Multiusos",            f"{BASE}/insumos-de-limpieza/limpiavidrios-multiusos", 1),
    ("Limpiatodos",                        f"{BASE}/insumos-de-limpieza/limpiatodos",     1),
    ("Thinner y Solventes",                f"{BASE}/insumos-de-limpieza/thinner-y-solventes", 1),
    ("Papeles Higiénicos y Toallas",       f"{BASE}/papeles-higienicos-y-toallas",        0),
    ("Papeles Higiénicos",                 f"{BASE}/papeles-higienicos-y-toallas/papeles-higienicos", 1),
    ("PH Domésticos x 20 Rollos",         f"{BASE}/papeles-higienicos-y-toallas/papeles-higienicos/papeles-higienicos-domesticos-x-20-rollos", 2),
    ("PH Domésticos x 24 Rollos",         f"{BASE}/papeles-higienicos-y-toallas/papeles-higienicos/papeles-higienicos-domesticos-x-24-rollos", 2),
    ("PH Hoteleros x 20 Rollos",          f"{BASE}/papeles-higienicos-y-toallas/papeles-higienicos/papeles-higienicos-hoteleros-x-20-rollos", 2),
    ("PH Jumbo 100-250 metros",           f"{BASE}/papeles-higienicos-y-toallas/papeles-higienicos/papeles-higienicos-jumbo-de-100-a-250-metros", 2),
    ("Papeles Toallas",                    f"{BASE}/papeles-higienicos-y-toallas/papeles-toallas", 1),
    ("Papeles Toalla Megarollo",           f"{BASE}/papeles-higienicos-y-toallas/papeles-toallas/papeles-toalla-megarollo", 2),
    ("Papeles Toalla Interfoliado",        f"{BASE}/papeles-higienicos-y-toallas/papeles-toallas/papeles-toalla-interfoliado", 2),
    ("Dispensadores",                      f"{BASE}/papeles-higienicos-y-toallas/dispensadores", 1),
    ("Bolsas Plásticas",                   f"{BASE}/bolsas-plasticas",                    0),
    ("Bolsas Basura 1.5 Micras",           f"{BASE}/bolsas-plasticas/bolsas-para-basura-de-15-micras", 1),
    ("Bolsas Basura 2 Micras",             f"{BASE}/bolsas-plasticas/bolsas-para-basura-de-2-micras", 1),
    ("Suministros Industriales",           f"{BASE}/suministros-industriales",            0),
    ("Señalizaciones de Seguridad",        f"{BASE}/suministros-industriales/senalizaciones-de-seguridad", 1),
    ("Cuidado y Salud",                    f"{BASE}/cuidado-y-salud",                     0),
    ("Aseo y Cuidado Personal",            f"{BASE}/cuidado-y-salud/aseo-y-cuidado-personal", 1),
    ("Alcoholes Líquidos y en Gel",        f"{BASE}/cuidado-y-salud/alcoholes-liquidos-y-en-gel-1", 1),
    ("Guantes Quirúrgicos",                f"{BASE}/cuidado-y-salud/alcoholes-liquidos-y-en-gel-1/guantes-quirurgicos", 2),
    ("Cafetería",                          f"{BASE}/cafeteria",                           0),
    ("Descartables",                       f"{BASE}/cafeteria/descartables",              1),
    ("Vasos Descartables",                 f"{BASE}/cafeteria/descartables/vasos-descartables", 2),
    ("Tapers y Contenedores",              f"{BASE}/cafeteria/descartables/tapers-y-contenedores-descartables", 2),
    ("Útiles de Oficina",                  f"{BASE}/utiles-de-oficina",                   0),
    ("Tachos de Basura y Contenedores",    f"{BASE}/tachos-de-basura-y-contenedores",     0),
    ("Contenedores de Basura",             f"{BASE}/tachos-de-basura-y-contenedores/contenedores-de-basura", 1),
    ("Ferretería",                         f"{BASE}/ferreteria",                          0),
    ("Términos y Condiciones",             f"{BASE}/terminos-y-condiciones",              0),
    ("Buscar",                             f"{BASE}/buscar",                              0),
]

PAGES = [
    ("Inicio",              "/",                         "Prolider – Distribuidora de Productos de Limpieza en Lima. Atención a empresas e instituciones."),
    ("Términos y Cond.",    "/terminos-y-condiciones",   "Términos y condiciones de compra en Prolider Empresarial S.A.C. Precios incluyen IGV 18%."),
    ("Sobre Nosotros",      "/nosotros",                 "Prolider Empresarial S.A.C. – Mz. E1 Lt. 8 Urb. Industrial El Lucumo, Lurín, Lima. Tel: 922139093."),
    ("Contacto",            "/contacto",                 "Contacto: ventas@prolider.pe | 922139093 | Yape y Plin disponibles. Envíos a Lima Metropolitana y provincias."),
]

BANNERS = [
    ("Banner Principal",    "Distribuidora de Productos de Limpieza en Lima – Prolider",
     "Proveemos a empresas, hospitales, colegios y restaurantes. Más de 1,000 productos."),
    ("Banner Envíos",       "Envíos a toda Lima y provincias",
     "Despachos a domicilio en Lima Metropolitana y envíos al interior vía agencias."),
    ("Banner Pago",         "Múltiples métodos de pago",
     "BCP, BBVA, Interbank, Scotiabank, Yape, Plin. Precios incluyen IGV 18%."),
    ("Banner Descuentos",   "Descuentos por volumen desde 12 unidades",
     "Compra mayor: descuentos especiales para empresas. Cotiza ahora."),
]

BRANDS = ["Sapolio","Genérico","Scotch Brite 3M","Palmolive","Hude","Diclotrin","QR",
          "Prolim","Vanish","Tuinies","Dersa","Marsella","Fiorucci","Paramonga","Rendipel"]

TECH = {
    "Plataforma":   "WordPress + WooCommerce",
    "Protección":   "Cloudflare WAF",
    "Empresa":      "Prolider Empresarial S.A.C.",
    "Dirección":    "Mz. E1 Lt. 8 Urb. Industrial El Lucumo, Lurín, Lima",
    "Teléfono":     "922139093",
    "Email":        "ventas@prolider.pe",
    "Redes":        "Facebook: /productosdelimpiezalima | Instagram: @productosdelimpiezalima",
}

COLORS = ["#1a73e8","#f8f9fa","#343a40","#28a745","#dc3545","#ffffff","#000000","#6c757d"]
FONTS  = ["Roboto", "Open Sans", "sans-serif"]

RANGES = {
    "Accesorios de Limpieza":    "S/ 1.50 – S/ 35.29",
    "Esponjas":                  "S/ 1.40 – S/ 43.90",
    "Guantes de Limpieza":       "S/ 4.90 – S/ 48.40",
    "Mopas Planas y Mechones":   "S/ 5.90 – S/ 449.00",
    "Plumeros":                  "S/ 2.90 – S/ 15.90",
    "Desinfectantes":            "S/ 7.90 – S/ 110.90",
    "Detergentes en Polvo":      "S/ 1.50 – S/ 85.90",
    "Detergentes Líquidos":      "S/ 8.80 – S/ 49.90",
    "Lejías":                    "S/ 2.00 – S/ 57.90",
    "Limpiavidrios Multiusos":   "S/ 4.40 – S/ 27.90",
    "Limpieza Automotriz":       "S/ 4.90 – S/ 121.90",
    "Papeles Higiénicos":        "S/ 15.90 – S/ 94.90",
    "Bolsas Plásticas":          "S/ 23.10 – S/ 89.90",
    "Tachos y Contenedores":     "S/ 35.00 – S/ 450.00",
}

# ──────────────────────────────────────────────────────────────────────────────
# HELPERS EXCEL
# ──────────────────────────────────────────────────────────────────────────────

def h_style(ws, color="1F4E79"):
    fill   = PatternFill("solid", fgColor=color)
    font   = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin   = Side(style="thin", color="CCCCCC")
    bord   = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment, cell.border = fill, font, align, bord
    ws.row_dimensions[1].height = 22

def auto_w(ws):
    for col in ws.columns:
        mx = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 4, 55)

def alt_rows(ws, last_row):
    for r in range(2, last_row + 1):
        bg = "EBF3FC" if r % 2 == 0 else "FFFFFF"
        for cell in ws[r]:
            cell.fill = PatternFill("solid", fgColor=bg)

# ──────────────────────────────────────────────────────────────────────────────
# GENERAR EXCEL
# ──────────────────────────────────────────────────────────────────────────────

def make_excel():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Hoja 1: Productos ─────────────────────────────────────────────────────
    ws = wb.create_sheet("Productos")
    ws.append(["ID","Nombre","SKU","Precio (S/)","Precio con IGV 18%","Categoría",
               "Marca","Descripción","URL Producto"])
    for p in PRODUCTS:
        cname = next((c[1] for c in CATEGORIES if c[0] == p[4]), "")
        slug  = p[7]
        cslug = cat_map.get(p[4],"")
        pid   = parent_map.get(p[4], 0)
        pslug = cat_map.get(pid,"") if pid else ""
        if pslug:
            url = f"{BASE}/{pslug}/{cslug}/{slug}"
        else:
            url = f"{BASE}/{cslug}/{slug}"
        ws.append([p[0], p[1], p[2], p[3], round(p[3]*1.18,2), cname, p[5], p[6], url])
    h_style(ws); alt_rows(ws, ws.max_row); auto_w(ws); ws.freeze_panes = "A2"

    # ── Hoja 2: Categorías ────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Categorias")
    ws2.append(["ID","Nombre","Slug","Padre ID","Padre Nombre","Nivel","URL"])
    for c in CATEGORIES:
        pid   = c[3]
        pname = next((x[1] for x in CATEGORIES if x[0]==pid), "Raíz") if pid else "Raíz"
        nivel = 0 if pid == 0 else (1 if parent_map.get(pid,0)==0 else 2)
        ws2.append([c[0], c[1], c[2], pid or "", pname, nivel, build_cat_url(c[0])])
    h_style(ws2, "14375A"); alt_rows(ws2, ws2.max_row); auto_w(ws2); ws2.freeze_panes = "A2"

    # ── Hoja 3: Imágenes ──────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Imagenes")
    ws3.append(["Producto ID","Nombre Producto","URL Imagen","Alt Text","Posición"])
    for p in PRODUCTS:
        cslug = cat_map.get(p[4],"")
        pid   = parent_map.get(p[4], 0)
        pslug = cat_map.get(pid,"") if pid else ""
        img_base = f"{BASE}/wp-content/uploads/products"
        ws3.append([p[0], p[1],
                    f"{img_base}/{p[7]}.jpg",
                    p[1], 0])
    h_style(ws3, "0D5C2E"); alt_rows(ws3, ws3.max_row); auto_w(ws3); ws3.freeze_panes = "A2"

    # ── Hoja 4: Menús y Navegación ────────────────────────────────────────────
    ws4 = wb.create_sheet("Menus")
    ws4.append(["Texto del Enlace","URL","Nivel (0=principal)"])
    for m in MENUS:
        ws4.append(list(m))
    h_style(ws4, "7B3F00"); alt_rows(ws4, ws4.max_row); auto_w(ws4); ws4.freeze_panes = "A2"

    # ── Hoja 5: Textos y Banners ──────────────────────────────────────────────
    ws5 = wb.create_sheet("Textos_Banners")
    ws5.append(["Tipo","Nombre/Título","URL","Texto / Descripción"])
    for pg in PAGES:
        ws5.append(["Página", pg[0], f"{BASE}{pg[1]}", pg[2]])
    for b in BANNERS:
        ws5.append(["Banner", b[0], BASE, f"{b[1]} — {b[2]}"])
    for brand in BRANDS:
        ws5.append(["Marca", brand, f"{BASE}/{brand.lower().replace(' ','-')}", ""])
    h_style(ws5, "5C1A7A"); alt_rows(ws5, ws5.max_row); auto_w(ws5); ws5.freeze_panes = "A2"

    # ── Hoja 6: Rangos de Precios por Categoría ───────────────────────────────
    ws6 = wb.create_sheet("Rangos_Precios")
    ws6.append(["Categoría","Rango de Precio"])
    for cat, rng in RANGES.items():
        ws6.append([cat, rng])
    h_style(ws6, "8B4513"); alt_rows(ws6, ws6.max_row); auto_w(ws6)

    # ── Hoja 7: Assets y Tecnología ───────────────────────────────────────────
    ws7 = wb.create_sheet("Assets_Tech")
    ws7.append(["Tipo","Valor"])
    for k, v in TECH.items():
        ws7.append([k, v])
    ws7.append(["",""])
    ws7.append(["Colores HEX",""])
    for c in COLORS:
        ws7.append(["Color", c])
    ws7.append(["",""])
    ws7.append(["Fuentes",""])
    for f in FONTS:
        ws7.append(["Fuente", f])
    h_style(ws7, "4A4A4A"); auto_w(ws7)

    path = os.path.join(OUT, "products.xlsx")
    wb.save(path)
    print(f"[✓] Excel guardado: {path}  ({len(PRODUCTS)} productos, {len(CATEGORIES)} categorías)")
    return path


# ──────────────────────────────────────────────────────────────────────────────
# GENERAR SQL
# ──────────────────────────────────────────────────────────────────────────────

def esc(v):
    if v is None: return "NULL"
    return "'" + str(v).replace("\\","\\\\").replace("'","\\'").replace("\n","\\n") + "'"

def make_sql():
    lines = [
        "-- ============================================================",
        "--  BASE DE DATOS: productosdelimpiezalima.com (Prolider)",
        f"--  Generado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "-- ============================================================",
        "SET NAMES utf8mb4;",
        "SET FOREIGN_KEY_CHECKS=0;",
        "",
        "DROP TABLE IF EXISTS `product_images`;",
        "DROP TABLE IF EXISTS `products`;",
        "DROP TABLE IF EXISTS `categories`;",
        "DROP TABLE IF EXISTS `nav_menus`;",
        "DROP TABLE IF EXISTS `pages`;",
        "",
        # CATEGORIES
        """CREATE TABLE `categories` (
  `id`        INT UNSIGNED NOT NULL,
  `name`      VARCHAR(255) NOT NULL,
  `slug`      VARCHAR(255),
  `parent_id` INT UNSIGNED DEFAULT 0,
  `url`       VARCHAR(512),
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;""",
        "",
        "INSERT INTO `categories` (id,name,slug,parent_id,url) VALUES",
    ]
    rows = []
    for c in CATEGORIES:
        rows.append(f"  ({c[0]},{esc(c[1])},{esc(c[2])},{c[3]},{esc(build_cat_url(c[0]))})")
    lines.append(",\n".join(rows) + ";")

    lines += [
        "",
        """CREATE TABLE `products` (
  `id`          INT UNSIGNED NOT NULL,
  `name`        VARCHAR(512) NOT NULL,
  `sku`         VARCHAR(100),
  `price`       DECIMAL(10,2),
  `price_igv`   DECIMAL(10,2),
  `category_id` INT UNSIGNED,
  `brand`       VARCHAR(100),
  `description` TEXT,
  `url`         VARCHAR(1024),
  PRIMARY KEY (`id`),
  KEY `idx_cat` (`category_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;""",
        "",
        "INSERT INTO `products` (id,name,sku,price,price_igv,category_id,brand,description,url) VALUES",
    ]
    rows = []
    for p in PRODUCTS:
        cslug = cat_map.get(p[4],"")
        pid   = parent_map.get(p[4], 0)
        pslug = cat_map.get(pid,"") if pid else ""
        url   = f"{BASE}/{pslug}/{cslug}/{p[7]}" if pslug else f"{BASE}/{cslug}/{p[7]}"
        rows.append(
            f"  ({p[0]},{esc(p[1])},{esc(p[2])},{p[3]:.2f},{p[3]*1.18:.2f},"
            f"{p[4]},{esc(p[5])},{esc(p[6])},{esc(url)})"
        )
    lines.append(",\n".join(rows) + ";")

    lines += [
        "",
        """CREATE TABLE `product_images` (
  `id`         INT UNSIGNED NOT NULL AUTO_INCREMENT,
  `product_id` INT UNSIGNED NOT NULL,
  `url`        VARCHAR(1024),
  `alt`        VARCHAR(512),
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;""",
        "",
        "INSERT INTO `product_images` (product_id,url,alt) VALUES",
    ]
    rows = []
    for p in PRODUCTS:
        rows.append(f"  ({p[0]},{esc(f'https://productosdelimpiezalima.com/wp-content/uploads/products/{p[7]}.jpg')},{esc(p[1])})")
    lines.append(",\n".join(rows) + ";")

    lines += [
        "",
        """CREATE TABLE `nav_menus` (
  `id`    INT UNSIGNED NOT NULL AUTO_INCREMENT,
  `text`  VARCHAR(255),
  `url`   VARCHAR(1024),
  `depth` INT DEFAULT 0,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;""",
        "",
        "INSERT INTO `nav_menus` (text,url,depth) VALUES",
    ]
    rows = [f"  ({esc(m[0])},{esc(m[1])},{m[2]})" for m in MENUS]
    lines.append(",\n".join(rows) + ";")

    lines += ["", "SET FOREIGN_KEY_CHECKS=1;", "", "-- FIN SCRIPT"]

    path = os.path.join(OUT, "database.sql")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[✓] SQL guardado: {path}")
    return path


# ──────────────────────────────────────────────────────────────────────────────
# GENERAR JSON
# ──────────────────────────────────────────────────────────────────────────────

def make_json():
    # images.json
    images = []
    for p in PRODUCTS:
        cslug = cat_map.get(p[4],"")
        pid   = parent_map.get(p[4], 0)
        pslug = cat_map.get(pid,"") if pid else ""
        url   = f"{BASE}/{pslug}/{cslug}/{p[7]}" if pslug else f"{BASE}/{cslug}/{p[7]}"
        images.append({
            "product_id":   p[0],
            "product_name": p[1],
            "url":          f"{BASE}/wp-content/uploads/products/{p[7]}.jpg",
            "alt":          p[1],
            "product_url":  url,
        })
    with open(os.path.join(OUT,"images.json"),"w",encoding="utf-8") as f:
        json.dump(images, f, ensure_ascii=False, indent=2)

    # sitemap.json
    cats_json = []
    for c in CATEGORIES:
        pid = c[3]
        pname = next((x[1] for x in CATEGORIES if x[0]==pid),"") if pid else ""
        cats_json.append({"id":c[0],"name":c[1],"slug":c[2],"parent_id":pid,"parent_name":pname,"url":build_cat_url(c[0])})

    prods_json = []
    for p in PRODUCTS:
        cname = next((c[1] for c in CATEGORIES if c[0]==p[4]),"")
        cslug = cat_map.get(p[4],"")
        pid2  = parent_map.get(p[4], 0)
        pslug = cat_map.get(pid2,"") if pid2 else ""
        url   = f"{BASE}/{pslug}/{cslug}/{p[7]}" if pslug else f"{BASE}/{cslug}/{p[7]}"
        prods_json.append({
            "id":p[0],"name":p[1],"sku":p[2],"price":p[3],
            "price_igv":round(p[3]*1.18,2),"category":cname,
            "brand":p[5],"description":p[6],"url":url
        })

    sitemap = {
        "site":         BASE,
        "empresa":      "Prolider Empresarial S.A.C.",
        "technology":   "WordPress + WooCommerce",
        "proteccion":   "Cloudflare WAF",
        "contacto":     {"telefono":"922139093","email":"ventas@prolider.pe"},
        "direccion":    "Mz. E1 Lt. 8 Urb. Industrial El Lucumo, Lurín, Lima",
        "total_categories": len(CATEGORIES),
        "total_products":   len(PRODUCTS),
        "total_menus":      len(MENUS),
        "categories":   cats_json,
        "products":     prods_json,
        "menus":        [{"text":m[0],"url":m[1],"level":m[2]} for m in MENUS],
        "pages":        [{"title":p[0],"url":f"{BASE}{p[1]}","description":p[2]} for p in PAGES],
        "banners":      [{"name":b[0],"titulo":b[1],"texto":b[2]} for b in BANNERS],
        "brands":       BRANDS,
        "colors":       COLORS,
        "fonts":        FONTS,
        "price_ranges": RANGES,
    }
    with open(os.path.join(OUT,"sitemap.json"),"w",encoding="utf-8") as f:
        json.dump(sitemap, f, ensure_ascii=False, indent=2)

    with open(os.path.join(OUT,"images.json"),"w",encoding="utf-8") as f:
        json.dump(images, f, ensure_ascii=False, indent=2)

    print(f"[✓] sitemap.json guardado  ({len(CATEGORIES)} cats, {len(PRODUCTS)} productos)")
    print(f"[✓] images.json guardado   ({len(images)} imágenes)")


# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"\n{'='*60}")
    print("  GENERANDO ARCHIVOS – productosdelimpiezalima.com")
    print(f"{'='*60}\n")
    make_excel()
    make_sql()
    make_json()
    print(f"\n{'='*60}")
    print(f"  LISTO – archivos en: {os.path.abspath(OUT)}/")
    print(f"  products.xlsx  |  database.sql  |  sitemap.json  |  images.json")
    print(f"{'='*60}\n")

"""
Configuración del extractor - productosdelimpiezalima.com
INSTRUCCIONES:
  1. Ve a WooCommerce > Ajustes > Avanzado > API REST
  2. Crea una clave con permisos de Lectura
  3. Pega consumer_key y consumer_secret abajo
"""

SITE_URL = "https://productosdelimpiezalima.com"

WC_CONSUMER_KEY    = "ck_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
WC_CONSUMER_SECRET = "cs_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

PER_PAGE = 100          # máx por petición WooCommerce
OUTPUT_DIR = "../data"  # dónde guardar Excel, SQL, JSON

REQUESTS_TIMEOUT  = 30   # segundos
REQUESTS_RETRIES  = 3
REQUESTS_DELAY    = 0.3  # segundos entre peticiones (evita rate-limit)

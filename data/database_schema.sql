-- ============================================================
--  SCHEMA BASE DE DATOS – productosdelimpiezalima.com
--  Archivo generado por VICZUL Extractor
--  Ejecuta woo_extractor.py para obtener el archivo con datos
-- ============================================================
SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS=0;

-- ── Categorías ───────────────────────────────────────────────
DROP TABLE IF EXISTS `categories`;
CREATE TABLE `categories` (
  `id`          INT UNSIGNED NOT NULL,
  `name`        VARCHAR(255) NOT NULL,
  `slug`        VARCHAR(255),
  `parent_id`   INT UNSIGNED DEFAULT 0,
  `description` TEXT,
  `count`       INT DEFAULT 0,
  `url`         VARCHAR(512),
  `image_url`   VARCHAR(512),
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ── Productos ────────────────────────────────────────────────
DROP TABLE IF EXISTS `products`;
CREATE TABLE `products` (
  `id`                INT UNSIGNED NOT NULL,
  `name`              VARCHAR(512) NOT NULL,
  `sku`               VARCHAR(255),
  `type`              VARCHAR(50),
  `status`            VARCHAR(50),
  `price`             DECIMAL(10,2),
  `regular_price`     DECIMAL(10,2),
  `sale_price`        DECIMAL(10,2),
  `stock_status`      VARCHAR(50),
  `stock_quantity`    INT,
  `manage_stock`      TINYINT(1) DEFAULT 0,
  `categories`        TEXT,
  `tags`              TEXT,
  `short_description` TEXT,
  `description`       LONGTEXT,
  `main_image_url`    VARCHAR(1024),
  `total_images`      INT DEFAULT 0,
  `attributes`        TEXT,
  `weight`            VARCHAR(50),
  `dimensions`        VARCHAR(255),
  `url`               VARCHAR(1024),
  `date_created`      DATETIME,
  `date_modified`     DATETIME,
  PRIMARY KEY (`id`),
  FULLTEXT KEY `ft_name_desc` (`name`,`description`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ── Imágenes de productos ────────────────────────────────────
DROP TABLE IF EXISTS `product_images`;
CREATE TABLE `product_images` (
  `id`            INT UNSIGNED NOT NULL AUTO_INCREMENT,
  `product_id`    INT UNSIGNED NOT NULL,
  `product_name`  VARCHAR(512),
  `image_id`      INT UNSIGNED,
  `url`           VARCHAR(1024),
  `alt`           VARCHAR(512),
  `position`      INT DEFAULT 0,
  PRIMARY KEY (`id`),
  KEY `idx_product` (`product_id`),
  CONSTRAINT `fk_img_product` FOREIGN KEY (`product_id`) REFERENCES `products` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ── Variantes ────────────────────────────────────────────────
DROP TABLE IF EXISTS `product_variations`;
CREATE TABLE `product_variations` (
  `id`             INT UNSIGNED NOT NULL AUTO_INCREMENT,
  `product_id`     INT UNSIGNED NOT NULL,
  `variation_id`   INT UNSIGNED,
  `sku`            VARCHAR(255),
  `price`          DECIMAL(10,2),
  `regular_price`  DECIMAL(10,2),
  `sale_price`     DECIMAL(10,2),
  `stock_status`   VARCHAR(50),
  `stock_quantity` INT,
  `attributes`     TEXT,
  `image_url`      VARCHAR(1024),
  PRIMARY KEY (`id`),
  KEY `idx_product` (`product_id`),
  CONSTRAINT `fk_var_product` FOREIGN KEY (`product_id`) REFERENCES `products` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ── Menús de navegación ──────────────────────────────────────
DROP TABLE IF EXISTS `nav_menus`;
CREATE TABLE `nav_menus` (
  `id`      INT UNSIGNED NOT NULL AUTO_INCREMENT,
  `text`    VARCHAR(512),
  `url`     VARCHAR(1024),
  `classes` VARCHAR(512),
  `depth`   INT DEFAULT 0,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ── Páginas del sitio ────────────────────────────────────────
DROP TABLE IF EXISTS `pages`;
CREATE TABLE `pages` (
  `id`       INT UNSIGNED NOT NULL,
  `title`    VARCHAR(512),
  `slug`     VARCHAR(255),
  `url`      VARCHAR(1024),
  `excerpt`  TEXT,
  `content`  LONGTEXT,
  `status`   VARCHAR(50),
  `template` VARCHAR(255),
  `date`     DATETIME,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

SET FOREIGN_KEY_CHECKS=1;
-- ── FIN SCHEMA ───────────────────────────────────────────────

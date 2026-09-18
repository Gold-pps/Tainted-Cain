-- 合成宝袋配方库：建库 + 建表（可重复执行）
-- 用法：mysql -u root -p < schema.sql
--   或在 mysql 客户端里：SOURCE C:/Users/czy/Desktop/github-repositories/Tainted-Cain/schema.sql;

CREATE DATABASE IF NOT EXISTS isaac
  DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;

USE isaac;

-- 掉落物单价（来自 合成宝袋组件.xlsx，28 行）
CREATE TABLE IF NOT EXISTS ingredient (
  order_index INT PRIMARY KEY,          -- 1..28，与模组枚举一致
  name_cn     VARCHAR(32) NOT NULL UNIQUE,
  value       INT NOT NULL
) ENGINE=InnoDB;

-- 道具信息（来自 以撒的结合忏悔+_全道具信息表.xlsx，721 行）
CREATE TABLE IF NOT EXISTS item (
  item_id     INT PRIMARY KEY,
  name_en     VARCHAR(64),
  name_cn     VARCHAR(64),
  quality     INT,
  description VARCHAR(255),
  kind        VARCHAR(16),
  cain_rating INT,
  KEY idx_item_name (name_cn)
) ENGINE=InnoDB;

-- 核心表：一条配方记录
CREATE TABLE IF NOT EXISTS recipe_record (
  seed_str    VARCHAR(16) NOT NULL,
  seed_int    INT UNSIGNED,
  recipe      VARCHAR(128) NOT NULL,
  output_id   INT NOT NULL,
  value       INT NULL,
  crafted     TINYINT NULL,             -- 1=是 0=否 NULL=未知
  sacred_orb  TINYINT NULL,             -- 1=有 0=无 NULL=未知
  first_seen  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  last_update DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
              ON UPDATE CURRENT_TIMESTAMP,
  source      VARCHAR(255),
  PRIMARY KEY (seed_str, recipe, output_id),
  KEY idx_output (output_id),
  KEY idx_seed_int (seed_int),
  CONSTRAINT fk_record_item FOREIGN KEY (output_id) REFERENCES item(item_id)
) ENGINE=InnoDB;

-- 配方明细：方便按掉落物反查/统计
CREATE TABLE IF NOT EXISTS recipe_line (
  seed_str    VARCHAR(16) NOT NULL,
  recipe      VARCHAR(128) NOT NULL,
  output_id   INT NOT NULL,
  order_index INT NOT NULL,
  qty         TINYINT NOT NULL,
  PRIMARY KEY (seed_str, recipe, output_id, order_index),
  CONSTRAINT fk_line_ing FOREIGN KEY (order_index)
    REFERENCES ingredient(order_index)
) ENGINE=InnoDB;

-- 变更历史：否→是、圣球状态翻转
CREATE TABLE IF NOT EXISTS recipe_event (
  id         BIGINT AUTO_INCREMENT PRIMARY KEY,
  seed_str   VARCHAR(16),
  recipe     VARCHAR(128),
  output_id  INT,
  field      VARCHAR(32),
  old_value  VARCHAR(64),
  new_value  VARCHAR(64),
  ts         DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- 查询视图：直接带出产物中文名与品质
CREATE OR REPLACE VIEW v_recipe AS
SELECT r.*, i.name_cn AS output_name, i.quality, i.kind
FROM recipe_record r
LEFT JOIN item i ON i.item_id = r.output_id;

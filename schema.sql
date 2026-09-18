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
  cain_rating INT,
  KEY idx_item_name (name_cn)
) ENGINE=InnoDB;

-- 核心表：一条配方记录
CREATE TABLE IF NOT EXISTS recipe_record (
  seed_str    VARCHAR(16) NOT NULL,
  recipe      VARCHAR(128) NOT NULL,
  output_id   INT NOT NULL,
  value       INT NULL,
  crafted     TINYINT NULL,             -- 1=是 0=否 NULL=未知
  sacred_orb  TINYINT NULL,             -- 1=有 0=无 NULL=未知
  craft_count INT NULL,                 -- 该配方产出该道具时，在该局被合成的次数
  PRIMARY KEY (seed_str, recipe, output_id),
  KEY idx_output (output_id),
  CONSTRAINT fk_record_item FOREIGN KEY (output_id) REFERENCES item(item_id)
) ENGINE=InnoDB;

-- 配方明细：方便按掉落物反查/统计（派生数据，导入时按 recipe_record 全量重建）
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

-- 说明：不保留变更历史表（不需要追溯“何时变成有圣球”这类时间线）。
--       若要回溯历史状态，改用定期快照：每次导入前把源文件拷到 snapshots/，
--       再加上每日 mysqldump。

-- 查询视图：直接带出产物中文名与品质
CREATE OR REPLACE VIEW v_recipe AS
SELECT r.seed_str, r.recipe, r.output_id, r.value, r.crafted, r.sacred_orb,
       r.craft_count, i.name_cn AS output_name, i.quality
FROM recipe_record r
LEFT JOIN item i ON i.item_id = r.output_id;

-- 1) 视图精简：只留你想要的列
CREATE OR REPLACE VIEW v_recipe AS
SELECT r.seed_str, r.recipe, r.output_id, r.value, r.crafted, r.sacred_orb,
       i.name_cn AS output_name, i.quality
FROM recipe_record r
LEFT JOIN item i ON i.item_id = r.output_id;

-- 2) 删掉底层列（这几列来自两张不同的表）
ALTER TABLE recipe_record
  DROP COLUMN seed_int,
  DROP COLUMN first_seen,
  DROP COLUMN last_update,
  DROP COLUMN source;

ALTER TABLE item DROP COLUMN kind;

-- 3) 验证
SELECT TABLE_NAME, COLUMN_NAME, ORDINAL_POSITION
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = 'isaac' AND TABLE_NAME IN ('recipe_record', 'item')
ORDER BY TABLE_NAME, ORDINAL_POSITION;

SELECT * FROM v_recipe LIMIT 5;

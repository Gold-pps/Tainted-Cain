SELECT CONCAT('SELECT * FROM `', TABLE_NAME, '`;') AS 语句
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = 'isaac' ORDER BY TABLE_NAME;

SELECT * FROM isaac.ingredient;
SELECT * FROM isaac.item;
SELECT * FROM isaac.recipe_event;
SELECT * FROM isaac.recipe_line;
SELECT * FROM isaac.recipe_record;
SELECT * FROM isaac.v_recipe;

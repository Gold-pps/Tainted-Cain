SELECT seed_str
FROM recipe_record
GROUP BY seed_str
HAVING SUM(
    CASE
        WHEN output_id = (SELECT item_id FROM item WHERE name_cn = '死亡证明' LIMIT 1)
         AND crafted = 1
        THEN 1
        ELSE 0
    END
) = 0;
# -*- coding: utf-8 -*-
"""把模组存档 / 工具记录 / 参考表导入 MySQL（幂等，可反复运行）。

用法（PowerShell）：
    $env:ISAAC_DB_PASSWORD = "你的密码"
    uv run import_to_mysql.py

数据来源：
  1) <游戏目录>\\data\\crafting_recorder\\save*.dat   —— 模组存档，按槽位多个文件
     每行：配方|道具ID|中文名|种子|价值|是否合成|十字圣球
  2) <项目目录>\\combinations*.txt                    —— 工具自己的记录
     每行：配方|道具ID|中文名|价值|是否合成|十字圣球
  3) 合成宝袋组件.xlsx / 全道具信息表.xlsx             —— 参考表

注意：
  - 模组存档会被游戏整份重写，所以只当输入用；本脚本只 INSERT/UPDATE，从不删记录。
  - “是否合成”只从否升级为是；“价值/十字圣球”已知值不会被未知(NULL)覆盖。
"""
import glob
import os
import re
import sys

import pymysql
from openpyxl import load_workbook

PROJECT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT)
from seed import str2seed  # noqa: E402

DB = dict(
    host=os.environ.get("ISAAC_DB_HOST", "127.0.0.1"),
    port=int(os.environ.get("ISAAC_DB_PORT", "3306")),
    user=os.environ.get("ISAAC_DB_USER", "isaac"),
    password=os.environ.get("ISAAC_DB_PASSWORD", ""),   # 用环境变量，别写进仓库
    database=os.environ.get("ISAAC_DB_NAME", "isaac"),
    charset="utf8mb4",                                  # 中文必须 utf8mb4
)

GAME_ROOT = os.environ.get(
    "ISAAC_GAME_DIR",
    r"C:\Program Files (x86)\Steam\steamapps\common\The Binding of Isaac Rebirth")
GAME_DATA = os.path.join(GAME_ROOT, "data", "crafting_recorder")

INGREDIENTS_XLSX = os.path.join(PROJECT, "合成宝袋组件.xlsx")
PROPS_XLSX = os.path.join(PROJECT, "以撒的结合忏悔+_全道具信息表.xlsx")


def tri(s):
    """'是' -> 1, '否' -> 0, 其他/空 -> None(未知)"""
    return {"是": 1, "否": 0}.get((s or "").strip())


def to_int(s):
    try:
        return int(s)
    except (TypeError, ValueError):
        return None


def txt(v):
    """写进 recipe_event 的历史值：NULL 用 '-' 表示"""
    return "-" if v is None else str(v)


def load_ingredients():
    ws = load_workbook(INGREDIENTS_XLSX, read_only=True).active
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or r[0] is None or r[2] is None:
            continue
        out.append((int(r[2]), str(r[0]).strip(), to_int(r[1]) or 0))
    return out


def load_items():
    ws = load_workbook(PROPS_XLSX, read_only=True)["道具信息"]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or r[0] is None:
            continue
        pid = to_int(r[0])
        if pid is None:
            continue
        out.append((pid, r[1], r[2], to_int(r[3]), r[4], r[5], to_int(r[6])))
    return out


def record_sources():
    """[(路径, 是否模组存档, 文件名推断的种子)]"""
    src = [(p, True, "") for p in sorted(glob.glob(
        os.path.join(GAME_DATA, "save*.dat")))]
    for p in sorted(glob.glob(os.path.join(PROJECT, "combinations*.txt"))):
        base = os.path.basename(p)
        seed = "" if base == "combinations.txt" else \
            base[len("combinations_"):-len(".txt")].replace("_", " ")
        src.append((p, False, seed))
    return src


def refresh_reference_tables(cur):
    """参考表由 xlsx 完全决定，直接重建（结构简单、量小）"""
    ingredients = load_ingredients()
    items = load_items()

    cur.execute("SET FOREIGN_KEY_CHECKS=0")
    cur.execute("TRUNCATE TABLE recipe_line")
    cur.execute("TRUNCATE TABLE ingredient")
    cur.executemany(
        "INSERT INTO ingredient (order_index, name_cn, value) VALUES (%s,%s,%s)",
        ingredients)
    cur.execute("TRUNCATE TABLE item")
    cur.executemany(
        "INSERT INTO item (item_id, name_en, name_cn, quality, description,"
        " kind, cain_rating) VALUES (%s,%s,%s,%s,%s,%s,%s)", items)
    # 已有记录引用、但参考表里没有的 ID：补占位行，保证外键不断
    cur.execute(
        "INSERT IGNORE INTO item (item_id, name_cn)"
        " SELECT DISTINCT r.output_id, NULL FROM recipe_record r"
        " LEFT JOIN item i ON i.item_id = r.output_id"
        " WHERE i.item_id IS NULL")
    cur.execute("SET FOREIGN_KEY_CHECKS=1")

    print(f"参考表：掉落物 {len(ingredients)} 条，道具 {len(items)} 条")
    return {name: oi for oi, name, _ in ingredients}


def import_records(cur, order_by_name):
    added = updated = skipped = 0

    for path, is_dat, seed_hint in record_sources():
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            lines = [ln.strip() for ln in f if ln.strip()]

        for line in lines:
            parts = line.split("|")
            if len(parts) < 2:
                skipped += 1
                continue
            recipe = parts[0]
            output_id = to_int(parts[1])
            if output_id is None:
                skipped += 1
                continue

            if is_dat:   # 配方|ID|名字|种子|价值|是否合成|十字圣球
                seed = parts[3].strip().upper() if len(parts) > 3 else ""
                value = to_int(parts[4]) if len(parts) > 4 else None
                crafted = tri(parts[5]) if len(parts) > 5 else None
                orb = tri(parts[6]) if len(parts) > 6 else None
            else:        # 配方|ID|名字|价值|是否合成|十字圣球
                seed = seed_hint
                value = to_int(parts[3]) if len(parts) > 3 else None
                crafted = tri(parts[4]) if len(parts) > 4 else None
                orb = tri(parts[5]) if len(parts) > 5 else None

            try:
                seed_int = str2seed(seed)
            except Exception:
                seed_int = None

            # 参考表里没有这个道具 ID（模组道具/版本差异）时补占位行，避免外键报错
            raw_name = parts[2].strip() if len(parts) > 2 else ""
            if raw_name.startswith("#"):
                raw_name = ""
            cur.execute(
                "INSERT IGNORE INTO item (item_id, name_cn) VALUES (%s, %s)",
                (output_id, raw_name or None))

            cur.execute(
                "SELECT value, crafted, sacred_orb FROM recipe_record"
                " WHERE seed_str=%s AND recipe=%s AND output_id=%s",
                (seed, recipe, output_id))
            row = cur.fetchone()

            if row is None:
                cur.execute(
                    "INSERT INTO recipe_record (seed_str, seed_int, recipe,"
                    " output_id, value, crafted, sacred_orb, source)"
                    " VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                    (seed, seed_int, recipe, output_id, value, crafted, orb,
                     path))
                added += 1
            else:
                # 合并规则：已知值不被未知覆盖；crafted 只从 0 升到 1
                new_value = row[0] if row[0] is not None else value
                new_crafted = 1 if row[1] == 1 else crafted
                new_orb = row[2] if row[2] is not None else orb
                if (new_value, new_crafted, new_orb) != row:
                    cur.execute(
                        "UPDATE recipe_record SET value=%s, crafted=%s,"
                        " sacred_orb=%s, source=%s"
                        " WHERE seed_str=%s AND recipe=%s AND output_id=%s",
                        (new_value, new_crafted, new_orb, path,
                         seed, recipe, output_id))
                    for field, old, new in (("value", row[0], new_value),
                                            ("crafted", row[1], new_crafted),
                                            ("sacred_orb", row[2], new_orb)):
                        if old != new:
                            cur.execute(
                                "INSERT INTO recipe_event (seed_str, recipe,"
                                " output_id, field, old_value, new_value)"
                                " VALUES (%s,%s,%s,%s,%s,%s)",
                                (seed, recipe, output_id, field,
                                 txt(old), txt(new)))
                    updated += 1
                else:
                    skipped += 1

            # 明细行：按掉落物反查用（配方串如 红心1+硬币2+钥匙3）
            cur.execute(
                "DELETE FROM recipe_line WHERE seed_str=%s AND recipe=%s"
                " AND output_id=%s", (seed, recipe, output_id))
            for name, qty in re.findall(r"([^\d+|]+)(\d+)", recipe):
                order_index = order_by_name.get(name.strip())
                if order_index:
                    cur.execute(
                        "INSERT IGNORE INTO recipe_line (seed_str, recipe,"
                        " output_id, order_index, qty) VALUES (%s,%s,%s,%s,%s)",
                        (seed, recipe, output_id, order_index, int(qty)))

        cur.connection.commit()
        print(f"  已处理 {len(lines):>3} 行：{path}")

    return added, updated, skipped


def main():
    conn = pymysql.connect(**DB)
    cur = conn.cursor()

    order_by_name = refresh_reference_tables(cur)
    conn.commit()

    print("开始导入记录……")
    added, updated, skipped = import_records(cur, order_by_name)
    conn.commit()

    cur.execute("SELECT COUNT(*) FROM recipe_record")
    total = cur.fetchone()[0]
    print(f"导入完成：新增 {added} 条，更新 {updated} 条，"
          f"无变化 {skipped} 条；recipe_record 现有 {total} 条")

    conn.close()


if __name__ == "__main__":
    main()

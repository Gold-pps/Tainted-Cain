# -*- coding: utf-8 -*-
"""把模组存档 / 工具记录 / 参考表导入 MySQL（幂等，可反复运行）。

用法（PowerShell）：
    $env:ISAAC_DB_PASSWORD = "你的密码"
    uv run import_to_mysql.py

数据来源：
  1) <游戏目录>\\data\\crafting_recorder\\save*.dat   —— 模组存档，按槽位多个文件
     每行：配方|道具ID|中文名|种子|价值|是否合成|十字圣球|本局合成次数
  2) <项目目录>\\combinations*.txt                    —— 工具自己的记录
     每行：配方|道具ID|中文名|价值|是否合成|十字圣球|本局合成次数
  3) 合成宝袋组件.xlsx / 全道具信息表.xlsx             —— 参考表

注意：
  - 模组存档会被游戏整份重写，所以只当输入用；本脚本只 INSERT/UPDATE，从不删记录。
  - “是否合成”只从否升级为是；“价值/十字圣球”已知值不会被未知(NULL)覆盖；
    “本局合成次数”只增不减。
"""
import glob
import os
import re

import pymysql
from openpyxl import load_workbook

PROJECT = os.path.dirname(os.path.abspath(__file__))

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


def normalize_name(s):
    """掉落物/道具名归一化：去掉空白，用于跨文件、跨表格匹配"""
    return re.sub(r"\s+", "", (s or "").strip())


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
        # 只取 ID / 英文名 / 中文名 / 品质 / 介绍 / 里该隐评级（不再导入 主动被动）
        out.append((pid, r[1], r[2], to_int(r[3]), r[4], to_int(r[6])))
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


def ensure_schema(cur):
    """补齐脚本需要的列（幂等），避免因表结构落后而报 1054。"""
    cur.execute(
        "SELECT COUNT(*) FROM information_schema.COLUMNS"
        " WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'recipe_record'"
        "   AND COLUMN_NAME = 'craft_count'")
    if cur.fetchone()[0] == 0:
        cur.execute("ALTER TABLE recipe_record"
                    " ADD COLUMN craft_count INT NULL AFTER sacred_orb")
        print("表结构缺少 craft_count 列，已自动添加")


def refresh_reference_tables(cur):
    """参考表由 xlsx 完全决定，直接重建（结构简单、量小）"""
    ingredients = load_ingredients()
    items = load_items()

    cur.execute("SET FOREIGN_KEY_CHECKS=0")
    cur.execute("TRUNCATE TABLE ingredient")
    cur.executemany(
        "INSERT INTO ingredient (order_index, name_cn, value) VALUES (%s,%s,%s)",
        ingredients)
    cur.execute("TRUNCATE TABLE item")
    cur.executemany(
        "INSERT INTO item (item_id, name_en, name_cn, quality, description,"
        " cain_rating) VALUES (%s,%s,%s,%s,%s,%s)", items)
    # 已有记录引用、但参考表里没有的 ID：补占位行，保证外键不断
    cur.execute(
        "INSERT IGNORE INTO item (item_id, name_cn)"
        " SELECT DISTINCT r.output_id, NULL FROM recipe_record r"
        " LEFT JOIN item i ON i.item_id = r.output_id"
        " WHERE i.item_id IS NULL")
    cur.execute("SET FOREIGN_KEY_CHECKS=1")

    print(f"参考表：掉落物 {len(ingredients)} 条，道具 {len(items)} 条")
    # 名字归一化后建索引，避免因空格/大小写差异匹配不上
    return {normalize_name(name): oi for oi, name, _ in ingredients}


def import_records(cur):
    """把各来源的配方记录 upsert 进 recipe_record"""
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

            if is_dat:   # 配方|ID|名字|种子|价值|是否合成|十字圣球|本局合成次数
                seed = parts[3].strip().upper() if len(parts) > 3 else ""
                value = to_int(parts[4]) if len(parts) > 4 else None
                crafted = tri(parts[5]) if len(parts) > 5 else None
                orb = tri(parts[6]) if len(parts) > 6 else None
                count = to_int(parts[7]) if len(parts) > 7 else None
            else:        # 配方|ID|名字|价值|是否合成|十字圣球|本局合成次数
                seed = seed_hint
                value = to_int(parts[3]) if len(parts) > 3 else None
                crafted = tri(parts[4]) if len(parts) > 4 else None
                orb = tri(parts[5]) if len(parts) > 5 else None
                count = to_int(parts[6]) if len(parts) > 6 else None
            if count is None:
                # 旧记录没有次数栏：已合成按 1 兜底，明确没合成的按 0
                if crafted == 1:
                    count = 1
                elif crafted == 0:
                    count = 0

            # 参考表里没有这个道具 ID（模组道具/版本差异）时补占位行，避免外键报错
            raw_name = parts[2].strip() if len(parts) > 2 else ""
            if raw_name.startswith("#"):
                raw_name = ""
            cur.execute(
                "INSERT IGNORE INTO item (item_id, name_cn) VALUES (%s, %s)",
                (output_id, raw_name or None))

            cur.execute(
                "SELECT value, crafted, sacred_orb, craft_count FROM recipe_record"
                " WHERE seed_str=%s AND recipe=%s AND output_id=%s",
                (seed, recipe, output_id))
            row = cur.fetchone()

            if row is None:
                cur.execute(
                    "INSERT INTO recipe_record (seed_str, recipe, output_id,"
                    " value, crafted, sacred_orb, craft_count)"
                    " VALUES (%s,%s,%s,%s,%s,%s,%s)",
                    (seed, recipe, output_id, value, crafted, orb, count))
                added += 1
            else:
                # 合并规则：已知值不被未知覆盖；crafted 只从 0 升到 1；
                # craft_count 只增不减
                new_value = row[0] if row[0] is not None else value
                new_crafted = 1 if row[1] == 1 else crafted
                new_orb = row[2] if row[2] is not None else orb
                if row[3] is None and count is None:
                    new_count = None          # 两边都不知道，保持未知
                else:
                    new_count = max(row[3] or 0, count or 0)
                if (new_value, new_crafted, new_orb, new_count) != row:
                    cur.execute(
                        "UPDATE recipe_record SET value=%s, crafted=%s,"
                        " sacred_orb=%s, craft_count=%s"
                        " WHERE seed_str=%s AND recipe=%s AND output_id=%s",
                        (new_value, new_crafted, new_orb, new_count,
                         seed, recipe, output_id))
                    updated += 1
                else:
                    skipped += 1

        cur.connection.commit()
        print(f"  已处理 {len(lines):>3} 行：{path}")

    return added, updated, skipped


def rebuild_recipe_lines(cur, order_by_name):
    """按 recipe_record 全量重建配方明细（recipe_line 是派生数据）。

    从库里现有的记录反推，而不是只看本次来源文件里出现过哪些记录 ——
    这样即使某些记录已从模组存档里消失，明细也不会跟着丢。
    """
    cur.execute("SELECT seed_str, recipe, output_id FROM recipe_record")
    rows = cur.fetchall()
    cur.execute("TRUNCATE TABLE recipe_line")

    inserted = 0
    unknown = {}
    for seed_str, recipe, output_id in rows:
        for name, qty in re.findall(r"([^\d+|]+)(\d+)", recipe):
            order_index = order_by_name.get(normalize_name(name))
            if order_index is None:
                unknown[name.strip()] = unknown.get(name.strip(), 0) + 1
                continue
            cur.execute(
                "INSERT IGNORE INTO recipe_line (seed_str, recipe, output_id,"
                " order_index, qty) VALUES (%s,%s,%s,%s,%s)",
                (seed_str, recipe, output_id, order_index, int(qty)))
            inserted += 1

    if unknown:
        detail = "、".join(f"{k}({v})" for k, v in sorted(unknown.items()))
        print(f"注意：有 {len(unknown)} 个掉落物名在参考表里找不到，"
              f"对应明细已跳过：{detail}")
        print("   请检查 合成宝袋组件.xlsx 的“名称”列与模组的 PICKUP_NAMES 是否一致")
    return inserted


def main():
    conn = pymysql.connect(**DB)
    cur = conn.cursor()

    ensure_schema(cur)
    order_by_name = refresh_reference_tables(cur)
    conn.commit()

    print("开始导入记录……")
    added, updated, skipped = import_records(cur)
    conn.commit()

    print("重建配方明细……")
    line_count = rebuild_recipe_lines(cur, order_by_name)
    conn.commit()

    cur.execute("SELECT COUNT(*) FROM recipe_record")
    total = cur.fetchone()[0]
    print(f"导入完成：新增 {added} 条，更新 {updated} 条，"
          f"无变化 {skipped} 条；recipe_record 现有 {total} 条，"
          f"recipe_line {line_count} 条")

    conn.close()


if __name__ == "__main__":
    main()

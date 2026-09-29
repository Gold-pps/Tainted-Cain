# -*- coding: utf-8 -*-
"""把数据库内容导出为 Excel（.xlsx）。

用法：
    uv run export_excel.py
        # 导出 v_recipe / recipe_record / recipe_line / item / ingredient
        # 到项目下的「导出」目录

    uv run export_excel.py v_recipe recipe_record
        # 只导出指定的表/视图

    uv run export_excel.py --sql "SELECT * FROM v_recipe WHERE sacred_orb = 1" --out 有圣球的配方.xlsx
        # 导出任意查询结果

连接参数与导入脚本一致（环境变量 ISAAC_DB_HOST/PORT/USER/PASSWORD/NAME）。
"""
import os
import sys

import pandas as pd
import pymysql

DB = dict(
    host=os.environ.get("ISAAC_DB_HOST", "127.0.0.1"),
    port=int(os.environ.get("ISAAC_DB_PORT", "3306")),
    user=os.environ.get("ISAAC_DB_USER", "isaac"),
    password=os.environ.get("ISAAC_DB_PASSWORD", ""),
    database=os.environ.get("ISAAC_DB_NAME", "isaac"),
    charset="utf8mb4",
)

OUT_DIR = "导出"
DEFAULT_TARGETS = ["v_recipe", "recipe_record", "recipe_line", "item", "ingredient"]


def tri_label(value):
    """1/0/NULL -> 是/否/未知"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "未知"
    try:
        return "是" if int(value) == 1 else "否"
    except (TypeError, ValueError):
        return "未知"


def polish(df):
    """给 1/0 的字段补一列中文，方便在 Excel 里直接看"""
    for col, label in (("crafted", "是否合成"), ("sacred_orb", "十字圣球")):
        if col in df.columns and label not in df.columns:
            df[label] = df[col].map(tri_label)
    return df


def fetch(conn, sql):
    cur = conn.cursor()
    cur.execute(sql)
    cols = [d[0] for d in cur.description]
    rows = cur.fetchall()
    cur.close()
    return pd.DataFrame(list(rows), columns=cols)


def save(df, path, sheet="data"):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]
        ws.freeze_panes = "A2"                     # 冻结表头
        if len(df.columns):
            ws.auto_filter.ref = ws.dimensions     # 打开筛选
    print(f"  已导出 {len(df):>5} 行 x {len(df.columns)} 列 -> {path}")


def main():
    args = sys.argv[1:]

    sql = out = None
    if "--sql" in args:
        i = args.index("--sql")
        sql = args[i + 1] if len(args) > i + 1 else ""
        args = args[:i] + args[i + 2:]
    if "--out" in args:
        i = args.index("--out")
        out = args[i + 1] if len(args) > i + 1 else None
        args = args[:i] + args[i + 2:]

    targets = args or DEFAULT_TARGETS
    os.makedirs(OUT_DIR, exist_ok=True)

    conn = pymysql.connect(**DB)
    try:
        if sql:
            path = out or os.path.join(OUT_DIR, "query.xlsx")
            print("正在执行查询并导出……")
            save(polish(fetch(conn, sql)), path, "query")
        else:
            print("正在导出……")
            for name in targets:
                df = polish(fetch(conn, f"SELECT * FROM `{name}`"))
                save(df, os.path.join(OUT_DIR, f"{name}.xlsx"), name[:31])
    finally:
        conn.close()

    print(f"\n完成，文件在：{os.path.abspath(OUT_DIR)}")


if __name__ == "__main__":
    main()

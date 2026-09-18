# Tainted-Cain（里该隐合成宝袋配方记录工具）

记录《以撒的结合：忏悔+》里该隐的**合成宝袋（Bag of Crafting）**配方与产物，并把数据落到 MySQL，
用于研究**配方 → 产物**的对应关系，尤其是**十字圣球（Sacred Orb）改变产物**的现象。

- 游戏侧：REPENTOGON 模组自动采集，无需手抄
- 桌面侧：Tkinter 图形记录器（手工录入 / 检索 / 从模组存档同步）
- 数据侧：一键导入 MySQL，用 SQL 做统计与对照

---

## 整体架构（数据流）

```
游戏（REPENTOGON）
  └─ mod/main.lua                     每帧检查合成袋，袋子满/合成时写记录
       └─ Isaac.SaveModData
            └─ <游戏目录>\data\crafting_recorder\save<槽位>.dat   ← 每个存档槽位一个文件
                        │
        ① 实时同步（可选） │
        combo_recorder.py ←┘            轮询所有槽位文件，按「配方+产物」合并去重
             └─ combinations_<种子>.txt  ← 工具自己的记录（游戏不会碰它）
                        │
        ② 批量导入（推荐每次玩完跑一次）
        import_to_mysql.py ←┘           同时读取两个 xlsx 参考表
             └─ MySQL：库 isaac
                  recipe_record（记录本体）
                  recipe_line  （配方成分，派生）
                  item / ingredient（参考表）
                  v_recipe     （视图：记录 + 产物中文名/品质）
```

要点：`save*.dat` 由游戏管理、会被整份重写，所以它只当**输入**；数据库才是长期存储。

---

## 文件说明

| 文件 | 作用 |
|---|---|
| `mod/main.lua` | REPENTOGON 模组：读取合成袋内容、判断合成事件、写入记录 |
| `mod/metadata.xml` | 模组元数据（模组目录名 `crafting_recorder`） |
| `mod/item_names_zh.lua` | 自动生成：道具 ID → 中文名（721 条） |
| `combo_recorder.py` | 图形记录器：掉落物按钮（图标 + 数量）、种子校验、道具检索、查重、导入导出、模组存档同步 |
| `import_to_mysql.py` | 把模组存档 / 工具记录 / 参考表导入 MySQL（幂等，可反复运行） |
| `import_db.bat` | 双击即可导入（自动取密码、中文不乱码、结束前暂停） |
| `schema.sql` | 建库 + 建表 + 视图，可重复执行 |
| `seed.py` | 种子字符串 ↔ 32 位种子互转（移植自开源项目 IsaacDecraftingHuijiGadget） |
| `合成宝袋组件.xlsx` | 28 种掉落物：名称 / 品质（= 单价）/ 顺序（= 模组枚举值） |
| `以撒的结合忏悔+_全道具信息表.xlsx` | 721 个道具：ID / 英文名 / 中文名 / 品质 / 介绍 / 主动被动 / 里该隐评级 |
| `Crafting_ui_sprite.png` | 掉落物图标精灵图（16×16 网格） |
| `Crafting.md` | 需求演变与设计决策记录（开发日志） |
| `combinations*.txt` | 运行时生成：记录文件，每个种子一个 |
| `db_password.txt` | 可选：数据库密码（一行）。已被 `.gitignore` 的 `*.txt` 忽略 |
| `make_excel.py` | 由硬编码清单生成掉落物价值 xlsx |
| `check_sprite.py` / `debug_grid.py` | 精灵图尺寸与切图网格的调试工具 |
| `test_font.py` / `test_font2.py` | 跨平台字体可用性检测 |
| `组件.csv` | 早期版本的掉落物价值表（已被 xlsx 取代） |
| `src/tainted_cain/` | `uv init` 生成的占位包，未使用 |

---

## 快速开始

### 0. 环境要求

- Windows（脚本内置 Steam 默认安装路径，可用环境变量 `ISAAC_GAME_DIR` 覆盖）
- Python ≥ 3.13 + [uv](https://docs.astral.sh/uv/)
- 《以撒的结合：忏悔+》+ [REPENTOGON](https://repentogon.com/docs.html)
- MySQL 8（仅在需要数据库时）

```bash
uv sync          # 安装依赖：openpyxl / pillow / pymysql
```

### 1. 安装模组

把 `mod/` 下的三个文件复制到游戏模组目录，然后重启游戏：

```
<游戏目录>\mods\crafting_recorder\
    ├─ metadata.xml
    ├─ main.lua
    └─ item_names_zh.lua
```

Windows 默认路径示例：

```
C:\Program Files (x86)\Steam\steamapps\common\The Binding of Isaac Rebirth\mods\crafting_recorder\
```

模组不显示任何游戏内提示，记录写入：

```
<游戏目录>\data\crafting_recorder\save1.dat   （save1/2/3 对应存档槽位）
```

### 2. 运行记录器（可选的图形界面）

```bash
uv run combo_recorder.py
```

- 左键点掉落物 +1，右键 −1；总数必须正好 8 才能记录
- 输入种子后记录会写到 `combinations_<种子>.txt`
- 会自动同步模组存档里新增的记录（状态栏显示来源文件）

### 3. 导入 MySQL

```bash
# 1) 建库建表：先进客户端，再用 SOURCE 执行
mysql -u root -p
#   mysql> SOURCE C:/Users/czy/Desktop/github-repositories/Tainted-Cain/schema.sql;
#   （PowerShell 不支持 "<" 重定向；在 cmd.exe 里才能写 mysql -u root -p < schema.sql）

# 2) 建专用账号
#    CREATE USER 'isaac'@'localhost' IDENTIFIED BY '你的密码';
#    GRANT ALL PRIVILEGES ON isaac.* TO 'isaac'@'localhost';

# 3) 让导入脚本能拿到密码（二选一）
setx ISAAC_DB_PASSWORD "你的密码"      # 写进用户环境变量
# 或者在项目目录新建 db_password.txt，内容就一行密码
```

之后**双击 `import_db.bat`**（等价于 `uv run import_to_mysql.py`）。
建议每次玩完游戏跑一次 —— `save*.dat` 随时可能被游戏覆盖。

---

## 记录格式

模组存档（`.dat`）每行 8 栏：

```
配方|道具ID|中文名|种子|掉落物价值之和|是否合成|十字圣球|本局合成次数

硬币1+钥匙4+炸弹1+胶囊2|72|念珠|TVDR ETCR|15|否|0|0
红心1+魂心1+硬币2+钥匙1+炸弹1+微型电池1+卡牌1|152|科技II|K9LH 6QXT|15|是|否|2
```

| 字段 | 说明 |
|---|---|
| 配方 | 按掉落物枚举升序拼接，如 `红心1+硬币2+钥匙3` |
| 道具ID | 合成产物 ID（`GetBagOfCraftingOutput()`） |
| 中文名 | 由 `item_names_zh.lua` 翻译（游戏内 `ItemConfig.Name` 只给出 `#XXX_NAME` 语言键） |
| 种子 | `GetStartSeedString()`，形如 `K9LH 6QXT` |
| 掉落物价值之和 | Σ 数量 × 单价，单价见 `合成宝袋组件.xlsx` 的“品质”列 |
| 是否合成 | 凑满袋子先记“否”，挥袋合成后**原地**改“是” |
| 十字圣球 | 记录时是否持有圣球（持有会改变产物品质） |
| 本局合成次数 | 该配方产出该产物时，这一局被真正合成过几次（只凑满不算） |

工具自己的 `combinations_*.txt` 与数据库里的记录少一栏「种子」（种子由文件名或记录字段承载）。

---

## 数据库结构

| 对象 | 类型 | 说明 |
|---|---|---|
| `recipe_record` | 表 | 记录本体，一行 = 一条配方事实。主键 `(seed_str, recipe, output_id)` |
| `recipe_line` | 表 | **派生表**：把配方拆成「掉落物 + 数量」，供按掉落物聚合/反查；导入时按 `recipe_record` 全量重建 |
| `item` | 表 | 道具参考表（由 `以撒的结合忏悔+_全道具信息表.xlsx` 决定，每次导入重建） |
| `ingredient` | 表 | 掉落物参考表（由 `合成宝袋组件.xlsx` 决定） |
| `v_recipe` | 视图 | `recipe_record` + `item`，带 `output_name`（产物中文名）与 `quality` |

`recipe_record` 关键列：

```
seed_str    种子（如 TVDR ETCR）
recipe      配方（如 硬币1+钥匙4+炸弹1+胶囊2）
output_id   产物 ID
value       掉落物价值之和
crafted     1=已合成 0=只凑过 NULL=未知
sacred_orb  1=有圣球 0=无圣球 NULL=未知
craft_count 该配方产出该产物时、在该局被合成的次数
```

不保留变更历史表；需要回溯历史状态时用快照（见下文“已知限制”）。

---

## 常用 SQL

```sql
-- 各局进度
SELECT seed_str, COUNT(*) AS 条数,
       SUM(crafted = 1) AS 已合成, SUM(crafted = 0) AS 只凑过,
       SUM(sacred_orb = 1) AS 有圣球时记录
FROM recipe_record GROUP BY seed_str ORDER BY 条数 DESC;

-- 圣球是否改变了产物（同种子同配方、有无圣球各一次、产物不同）
SELECT a.seed_str, a.recipe,
       CONCAT(a.output_id, ' ', a.output_name, '（品质', a.quality, '）') AS 无圣球,
       CONCAT(b.output_id, ' ', b.output_name, '（品质', b.quality, '）') AS 有圣球
FROM v_recipe a JOIN v_recipe b
  ON a.seed_str = b.seed_str AND a.recipe = b.recipe
 AND a.sacred_orb = 0 AND b.sacred_orb = 1 AND a.output_id <> b.output_id;

-- 掉落物使用频次（需要 recipe_line）
SELECT i.name_cn AS 掉落物, SUM(l.qty) AS 总用量, COUNT(*) AS 出现在几个配方里
FROM recipe_line l JOIN ingredient i ON i.order_index = l.order_index
GROUP BY i.name_cn ORDER BY 总用量 DESC;

-- 某个掉落物参与过哪些配方（精确匹配，避免 LIKE 撞子串）
SELECT i.name_cn AS 掉落物, l.qty AS 用量,
       v.seed_str, v.recipe, v.output_name, v.quality, v.crafted, v.sacred_orb
FROM recipe_line l
JOIN ingredient i ON i.order_index = l.order_index
JOIN v_recipe v   ON v.seed_str = l.seed_str
                 AND v.recipe = l.recipe AND v.output_id = l.output_id
WHERE i.name_cn = '金钥匙'
ORDER BY v.quality DESC;

-- 只凑过、还没真合成的配方
SELECT seed_str, recipe, output_name, value, sacred_orb, craft_count
FROM v_recipe WHERE crafted = 0 ORDER BY seed_str, value DESC;
```

注意：`v_recipe.output_name` 是视图里的别名，`item` 表里对应列叫 `name_cn`；
用 `LIKE '%钥匙%'` 会同时命中「钥匙 / 金钥匙 / 充能钥匙 / 红钥匙碎片」，要精确筛选请用 `recipe_line`。

---

## 命令速查

| 目的 | 命令 |
|---|---|
| 安装依赖 | `uv sync` |
| 启动记录器 | `uv run combo_recorder.py` |
| 导入数据库 | 双击 `import_db.bat`，或 `uv run import_to_mysql.py` |
| 建库建表 | 进客户端后 `SOURCE <绝对路径>/schema.sql;`（PowerShell 不支持 `<` 重定向） |
| 进库查询 | `mysql -u isaac -p isaac` |
| 单条查询 | `mysql -u isaac -p isaac -e "SELECT COUNT(*) FROM recipe_record;"` |
| 导出整库备份 | `mysqldump -u isaac -p --result-file=isaac_dump.sql isaac` |

环境变量：`ISAAC_GAME_DIR`、`ISAAC_DB_HOST`、`ISAAC_DB_PORT`、`ISAAC_DB_USER`、`ISAAC_DB_PASSWORD`、`ISAAC_DB_NAME`。

---

## 已知限制与坑

1. **模组存档由游戏管理**：`data\crafting_recorder\save*.dat` 会被游戏按内存副本整份重写，
   手工编辑留不住；不同存档槽位各有一个文件、内容可能重叠，因此程序会**合并所有槽位文件**。
   数据库才是长期存储，建议每次玩完立刻导入一次。
2. **去重键是「种子 + 配方 + 产物」**：同一配方若因圣球改变产物，会分成两行，`craft_count` 各自独立计数。
3. **配方字符串用中文名**：模组 `PICKUP_NAMES` 必须与 `合成宝袋组件.xlsx` 的“名称”列一致。
   不一致时导入会提示「有 N 个掉落物名在参考表里找不到」，对应明细行会被跳过。
4. **种子靠文件名/记录字段推断**：把 `combinations_XXXX_XXXX.txt` 改名（如加“- 副本”）会导致归到错误种子。
5. **NULL 与 0 含义不同**：`crafted` / `sacred_orb` / `craft_count` 用 `NULL` 表示“未知”，不要用 0 代替。
6. **便便（枚举 29）不参与配方与价值计算**：若游戏允许它进合成袋，该配方会与游戏显示不一致。
7. **中文排序**：库用 `utf8mb4_general_ci`，只影响排序观感，不影响数据。
8. **`import_db.bat` 里不能写中文注释**：cmd 按 GBK 解析批处理文件，中文会乱码并破坏语法。
9. **Python 脚本的 `print` 不要用 emoji**：GBK 控制台会抛 `UnicodeEncodeError`（改成“注意：”这类纯文本）。
10. **历史状态无法回溯**：已移除变更历史表，改用「导入前把源文件快照到 `snapshots/`」+「每日 `mysqldump`」。

---

## 参考表更新

| 想改什么 | 改哪里 | 是否要动模组 |
|---|---|---|
| 掉落物单价 | `合成宝袋组件.xlsx` 的“品质”列 | 要（`main.lua` 里的 `PICKUP_VALUES`） |
| 掉落物名称/顺序 | `合成宝袋组件.xlsx` 的“名称/顺序”列 | 要（`main.lua` 里的 `PICKUP_NAMES`） |
| 道具名/品质/评级 | `以撒的结合忏悔+_全道具信息表.xlsx` | 要（`item_names_zh.lua` 需按表重新生成） |

改完 xlsx 后重跑一次导入，参考表会自动重建；模组相关改动需把 `main.lua` 重新复制到游戏目录并重启游戏。

---

## 相关文档

- `Crafting.md`：需求演变、字段口径与设计决策的完整记录
- REPENTOGON 文档：<https://repentogon.com/docs.html>
- 种子算法来源：IsaacDecraftingHuijiGadget（`seed.py` 顶部有说明）
- 图标资源取自游戏本体，仅供个人使用

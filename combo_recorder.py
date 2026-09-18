import os
import re
import glob
import tkinter as tk
from tkinter import messagebox, filedialog
from openpyxl import load_workbook
import tkinter.font as tkfont

from seed import str2seed, is_valid_seed

INGREDIENTS_FILE = "合成宝袋组件.xlsx"
PROPS_FILE = "以撒的结合忏悔+_全道具信息表.xlsx"

# 模组（crafting_recorder）用 Isaac.SaveModData 落盘的目录名
MOD_DATA_DIR_NAME = "crafting_recorder"

# ============ 精灵图配置 ============
SPRITE_SHEET = "Crafting_ui_sprite.png"
SPRITE_SIZE = 40

# 图标在 sprite sheet 中的位置 (x, y, 宽, 高)
# 顶部：从第一行第 2 格 (16,0) 开始，按从左到右、从上到下，共 26 个
# 顶部 28 个：从第一行第 2 格 (16,0) 开始，遍历 4 行，跳过 (0,0)
SPRITE_POSITIONS = []
_count = 0
for _y in range(0, 64, 16):        # 4 行
    for _x in range(0, 128, 16):   # 每行 8 格
        if _count == 0 and _x == 0 and _y == 0:
            continue               # 跳过第一行第一格
        SPRITE_POSITIONS.append((_x, _y, 16, 16))
        _count += 1
        if _count >= 28:
            break
    if _count >= 28:
        break

# 这些字体变量会在 __main__ 中根据系统可用字体动态赋值
FONT_NORMAL = None
FONT_BOLD = None
FONT_TITLE = None
FONT_SMALL = None
FONT_MONO = None
FONT_SEED = None


# ---------- 通用工具 ----------
def normalize_seed(s):
    s = s.strip().upper()
    if " " not in s and len(s) == 8:
        s = s[:4] + " " + s[4:]
    return s


def get_filename(seed):
    if seed:
        return f"combinations_{seed.replace(' ', '_')}.txt"
    return "combinations.txt"


def parse_record(line):
    parts = line.split("|")
    if len(parts) >= 3:
        return parts[0], parts[1], parts[2]
    return parts[0], "", ""


def game_root_candidates():
    """以撒游戏安装目录的候选路径。

    可通过环境变量 ISAAC_GAME_DIR 直接指定。
    """
    roots = []
    env = os.environ.get("ISAAC_GAME_DIR")
    if env:
        roots.append(env)

    game_name = "The Binding of Isaac Rebirth"
    for drive in ("C:", "D:", "E:", "F:", "G:"):
        roots.append(os.path.join(drive + os.sep, "Program Files (x86)",
                                  "Steam", "steamapps", "common", game_name))
        roots.append(os.path.join(drive + os.sep, "Program Files", "Steam",
                                  "steamapps", "common", game_name))
        roots.append(os.path.join(drive + os.sep, "SteamLibrary",
                                  "steamapps", "common", game_name))
        roots.append(os.path.join(drive + os.sep, "Steam", "steamapps",
                                  "common", game_name))
    return roots


def find_mod_saves():
    """查找 Crafting Recorder 模组的所有 ModData 存档（save1/2/3.dat）。

    模组用 Isaac.SaveModData 落盘，实际位置是游戏安装目录下的
    ``data\\crafting_recorder\\save<槽位>.dat``。
    不同存档槽位各有自己的文件、内容可能重叠，所以要把它们全部读出来合并；
    部分环境下也可能出现在「我的文档\\My Games」里，故两处都找。
    """
    found = []

    def scan(base):
        d = os.path.join(base, "data", MOD_DATA_DIR_NAME)
        if os.path.isdir(d):
            found.extend(glob.glob(os.path.join(d, "save*.dat")))

    for root in game_root_candidates():
        scan(root)

    docs = os.path.join(os.path.expanduser("~"), "Documents", "My Games")
    for game_dir in ("Binding of Isaac Repentance+",
                     "Binding of Isaac Repentance"):
        scan(os.path.join(docs, game_dir))

    if not found:
        return []
    # 去重后按写入时间从旧到新排列：新的状态后处理，可以覆盖旧的
    uniq = sorted({os.path.normcase(p) for p in found})
    return sorted(uniq, key=os.path.getmtime)


# ---------- 数据加载 ----------
def load_ingredients():
    wb = load_workbook(INGREDIENTS_FILE)
    ws = wb.active
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 3:
            continue
        name, quality, order = row[0], row[1], row[2]
        if name and order is not None:
            items.append((int(order), str(name).strip()))
    items.sort()
    return items


def load_ingredient_values():
    """掉落物名称 -> 单件价值（《合成宝袋组件.xlsx》的“品质”列）"""
    wb = load_workbook(INGREDIENTS_FILE, read_only=True)
    ws = wb.active
    values = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 3:
            continue
        name, quality, order = row[0], row[1], row[2]
        if name and order is not None:
            try:
                values[str(name).strip()] = int(quality)
            except (TypeError, ValueError):
                pass
    return values


def recipe_value(recipe, values):
    """按配方字符串（如 红心2+硬币4+钥匙2）算掉落物价值之和"""
    total = 0
    for name, qty in re.findall(r"([^\d+|]+)(\d+)", recipe):
        total += values.get(name.strip(), 0) * int(qty)
    return total


def load_props():
    wb = load_workbook(PROPS_FILE, read_only=True)
    ws = wb["道具信息"]
    props = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 3:
            continue
        pid, en, cn = row[0], row[1], row[2]
        try:
            pid_int = int(pid)
        except (ValueError, TypeError):
            continue
        props.append((pid_int, str(en or "").strip(), str(cn or "").strip()))
    props.sort()
    return props


def load_combinations(filename):
    if not os.path.exists(filename):
        return []
    with open(filename, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]


def save_combination(filename, record):
    with open(filename, "a", encoding="utf-8") as f:
        f.write(record + "\n")


# ---------- 精灵图切分 ----------
def load_sprites(sheet_path, item_names, size=40):
    """按 SPRITE_POSITIONS 切图"""
    from PIL import Image, ImageTk

    sprites = {}
    refs = []

    if not os.path.exists(sheet_path):
        print(f"⚠️ 找不到 {sheet_path}，将使用纯文字按钮")
        return sprites, refs

    try:
        sheet = Image.open(sheet_path).convert("RGBA")
    except Exception as e:
        print(f"⚠️ 打开 {sheet_path} 失败：{e}")
        return sprites, refs

    for i, name in enumerate(item_names):
        if i >= len(SPRITE_POSITIONS):
            break
        x, y, cw, ch = SPRITE_POSITIONS[i]
        img = sheet.crop((x, y, x + cw, y + ch))
        img = img.resize((size, size), Image.NEAREST)

        photo = ImageTk.PhotoImage(img)
        sprites[name] = photo
        refs.append(photo)

    return sprites, refs

# ---------- 主界面 ----------
class ComboRecorder:
    def __init__(self, root, ingredients, props, sprites):
        self.root = root
        self.root.title("合成宝袋配方记录器")
        self.root.geometry("720x960")
        self.root.configure(bg="#FAFAFA")
        self.root.resizable(True, True)

        self.ingredients = ingredients
        self.props = props
        self.sprites = sprites
        self.filtered_props = list(props)
        self.quantities = {name: 0 for _, name in ingredients}
        self.selected_prop = None
        self.seed = ""
        self.filename = get_filename("")
        self.combos = load_combinations(self.filename)

        # 道具 ID -> 中文名，用于同步模组数据时补全记录
        self.prop_id_to_name = {pid: cn for pid, _, cn in props}
        # 掉落物名称 -> 单件价值，用于补全/计算“掉落物价值之和”
        self.ingredient_values = load_ingredient_values()

        # 模组自动同步状态
        self._sync_paths = find_mod_saves()   # 各存档槽位的 save*.dat（可能多个）
        self._sync_sigs = {}    # 路径 -> (大小, 修改时间) 指纹，用于判断是否有更新
        self._sync_total = 0
        self._sync_retry = 0
        self._sync_status_text = ""
        self._sync_cache = {}  # 非当前种子的记录文件 -> 行列表

        root.grid_columnconfigure(0, weight=1)

        # ============ 种子输入 ============
        seed_frame = tk.Frame(root, bg="#FAFAFA")
        seed_frame.grid(row=0, column=0, pady=(8, 2))

        tk.Label(seed_frame, text="种子：", font=FONT_NORMAL,
                 bg="#FAFAFA").pack(side="left")
        self.seed_entry = tk.Entry(seed_frame, font=FONT_SEED,
                                   width=12, justify="center")
        self.seed_entry.pack(side="left", padx=6)
        self.seed_entry.bind("<Return>", lambda e: self.apply_seed())

        for txt, cmd in [("应用", self.apply_seed), ("清除", self.clear_seed)]:
            tk.Button(seed_frame, text=txt, width=5,
                      font=FONT_SMALL,
                      bg="#FFFFFF", relief="solid", bd=1, cursor="hand2",
                      command=cmd).pack(side="left", padx=2)

        self.seed_status = tk.Label(root, text="当前无种子（记录到 combinations.txt）",
                                    font=FONT_SMALL,
                                    bg="#FAFAFA", fg="#616161")
        self.seed_status.grid(row=1, column=0, pady=(0, 4))

        # ============ 顶部状态 ============
        top = tk.Frame(root, bg="#FAFAFA")
        top.grid(row=2, column=0, pady=(2, 4))

        self.total_label = tk.Label(top, text="总数：0 / 8",
                                    font=FONT_TITLE,
                                    bg="#FAFAFA", fg="#212121")
        self.total_label.pack()

        self.recipe_label = tk.Label(top, text="当前配方：（空）",
                                     font=FONT_NORMAL,
                                     bg="#FAFAFA", fg="#1976D2",
                                     wraplength=660, justify="center")
        self.recipe_label.pack(pady=(4, 0))

        # ============ 物品按钮（带图标） ============
        btn_frame = tk.Frame(root, bg="#FAFAFA")
        btn_frame.grid(row=3, column=0, pady=2)

        self.buttons = {}
        COLS = 4
        for i, (order, name) in enumerate(ingredients):
            icon = sprites.get(name)

            if icon:
                btn = tk.Button(
                    btn_frame,
                    image=icon,
                    text="0",
                    compound="top",
                    width=76, height=72,
                    font=FONT_NORMAL,
                    bg="#FFFFFF", activebackground="#E3F2FD",
                    relief="solid", bd=1, cursor="hand2",
                    command=lambda n=name: self.add_item(n),
                )
            else:
                btn = tk.Button(
                    btn_frame,
                    text="0",
                    width=6, height=3,
                    font=FONT_NORMAL,
                    bg="#FFFFFF", activebackground="#E3F2FD",
                    relief="solid", bd=1, cursor="hand2",
                    command=lambda n=name: self.add_item(n),
                )

            btn.bind("<Button-3>", lambda e, n=name: self.remove_item(n))
            btn.grid(row=i // COLS, column=i % COLS, padx=2, pady=2)
            self.buttons[name] = btn

        # ============ 道具检索 ============
        prop_header = tk.Frame(root, bg="#FAFAFA")
        prop_header.grid(row=4, column=0, pady=(8, 2), sticky="ew", padx=20)

        tk.Label(prop_header, text="道具检索：",
                 font=FONT_BOLD,
                 bg="#FAFAFA").pack(side="left")

        self.prop_search = tk.Entry(prop_header, font=FONT_NORMAL,
                                    width=22)
        self.prop_search.pack(side="left", padx=6)
        self.prop_search.bind("<KeyRelease>", self.on_search_change)
        self.prop_search.bind("<Return>", self.on_search_enter)
        self.prop_search.bind("<Escape>", lambda e: self.clear_prop_search())

        tk.Button(prop_header, text="清空检索", font=FONT_SMALL,
                  bg="#FFFFFF", relief="solid", bd=1, cursor="hand2",
                  command=self.clear_prop_search).pack(side="left", padx=2)

        tk.Button(prop_header, text="清除已选", font=FONT_SMALL,
                  bg="#FFFFFF", relief="solid", bd=1, cursor="hand2",
                  command=self.clear_prop_selection).pack(side="left", padx=2)

        info_frame = tk.Frame(root, bg="#FAFAFA")
        info_frame.grid(row=5, column=0, pady=(2, 2))

        self.match_label = tk.Label(info_frame, text=f"共 {len(self.props)} 个道具",
                                    font=FONT_SMALL,
                                    bg="#FAFAFA", fg="#616161")
        self.match_label.pack(side="left", padx=(0, 12))

        self.selected_label = tk.Label(info_frame, text="未选择道具",
                                       font=FONT_NORMAL,
                                       bg="#FAFAFA", fg="#9E9E9E")
        self.selected_label.pack(side="left")

        # 道具列表
        prop_list_frame = tk.Frame(root, bg="#FAFAFA")
        prop_list_frame.grid(row=6, column=0, padx=20, sticky="ew")
        prop_list_frame.grid_columnconfigure(0, weight=1)

        prop_scroll = tk.Scrollbar(prop_list_frame, orient="vertical")
        self.prop_listbox = tk.Listbox(prop_list_frame, height=5,
                                       font=FONT_MONO,
                                       yscrollcommand=prop_scroll.set,
                                       activestyle="none",
                                       bg="#FFFFFF", bd=1, relief="solid",
                                       highlightthickness=0,
                                       selectbackground="#BBDEFB",
                                       selectforeground="#000000")
        prop_scroll.config(command=self.prop_listbox.yview)
        self.prop_listbox.grid(row=0, column=0, sticky="ew")
        prop_scroll.grid(row=0, column=1, sticky="ns")
        self.prop_listbox.bind("<<ListboxSelect>>", self.on_prop_select)
        self.prop_listbox.bind("<Double-Button-1>", self.on_prop_double_click)
        self.prop_listbox.bind("<Return>", self.on_prop_select)

        self.refresh_prop_list()

        # ============ 操作按钮 ============
        bottom = tk.Frame(root, bg="#FAFAFA")
        bottom.grid(row=7, column=0, pady=6)

        for text, cmd in [
            ("记录配方", self.record_combo),
            ("清空当前", self.clear_current),
            ("删除记录文件", self.clear_records),
            ("导入记录", self.import_records),
            ("导出记录", self.export_records),
        ]:
            tk.Button(bottom, text=text, width=12,
                      font=FONT_NORMAL,
                      bg="#FFFFFF", activebackground="#E3F2FD",
                      relief="solid", bd=1, cursor="hand2",
                      command=cmd).pack(side="left", padx=5)

        # 手动记录时标记是否持有十字圣球（会改变合成袋产出的品质）
        self.orb_var = tk.BooleanVar(value=False)
        tk.Checkbutton(bottom, text="有十字圣球", variable=self.orb_var,
                       font=FONT_NORMAL, bg="#FAFAFA",
                       activebackground="#FAFAFA",
                       cursor="hand2").pack(side="left", padx=(8, 0))

        # ============ 已记录列表 ============
        list_frame = tk.Frame(root, bg="#FAFAFA")
        list_frame.grid(row=8, column=0, padx=20, pady=(0, 12), sticky="nsew")
        root.grid_rowconfigure(8, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        header = tk.Frame(list_frame, bg="#FAFAFA")
        header.grid(row=0, column=0, sticky="ew")

        tk.Label(header, text="已记录：",
                 font=FONT_BOLD,
                 bg="#FAFAFA", fg="#424242").pack(side="left")

        self.sync_label = tk.Label(header, text="",
                                   font=FONT_SMALL,
                                   bg="#FAFAFA", fg="#9E9E9E")
        self.sync_label.pack(side="right")

        inner = tk.Frame(list_frame, bg="#FAFAFA")
        inner.grid(row=1, column=0, sticky="nsew")
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_rowconfigure(0, weight=1)

        scrollbar = tk.Scrollbar(inner, orient="vertical")
        self.listbox = tk.Listbox(inner, font=FONT_MONO,
                                  yscrollcommand=scrollbar.set,
                                  activestyle="none",
                                  bg="#FFFFFF", bd=1, relief="solid",
                                  highlightthickness=0)
        scrollbar.config(command=self.listbox.yview)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.refresh_record_list()
        self.update_display()
        self._poll_mod_data()

    # ---------- 种子 ----------
    def apply_seed(self):
        raw = self.seed_entry.get().strip()
        if not raw:
            self.clear_seed()
            return
        s = normalize_seed(raw)
        if not is_valid_seed(s):
            messagebox.showerror("无效种子", f"种子 {s!r} 未通过校验。")
            return
        self._switch_seed(s)

    def clear_seed(self):
        self.seed_entry.delete(0, tk.END)
        self._switch_seed("")

    def _switch_seed(self, seed):
        self.seed = seed
        self.filename = get_filename(seed)
        self.combos = load_combinations(self.filename)
        self._sync_cache.pop(self.filename, None)

        if seed:
            self.seed_status.config(
                text=f"当前种子：{seed}  (0x{str2seed(seed):08X})",
                fg="#2E7D32")
        else:
            self.seed_status.config(
                text="当前无种子（记录到 combinations.txt）",
                fg="#616161")

        self.refresh_record_list()
        self.clear_current()

    # ---------- 配方操作 ----------
    def add_item(self, name):
        if sum(self.quantities.values()) >= 8:
            return
        self.quantities[name] += 1
        self.update_display()

    def remove_item(self, name):
        if self.quantities[name] > 0:
            self.quantities[name] -= 1
            self.update_display()

    def update_display(self):
        total = sum(self.quantities.values())
        self.total_label.config(text=f"总数：{total} / 8")

        for name, btn in self.buttons.items():
            btn.config(text=str(self.quantities[name]))

        parts = []
        for _, name in self.ingredients:
            qty = self.quantities[name]
            if qty > 0:
                parts.append(f"{name}{qty}")
        recipe = "+".join(parts) if parts else "（空）"
        self.recipe_label.config(text=f"当前配方：{recipe}", fg="#1976D2")

    def clear_current(self):
        for name in self.quantities:
            self.quantities[name] = 0
        self.clear_prop_selection()
        self.update_display()

    # ---------- 道具检索 ----------
    def on_search_change(self, event=None):
        if event and event.keysym in ("Return", "Escape", "Up", "Down",
                                      "Left", "Right", "Tab"):
            return

        query = self.prop_search.get().strip().lower()
        if not query:
            self.filtered_props = list(self.props)
        else:
            self.filtered_props = [
                p for p in self.props
                if query in p[2].lower()
                or query in p[1].lower()
                or query in str(p[0])
            ]

        self.refresh_prop_list()

        if query:
            exact = [p for p in self.filtered_props
                     if p[2].lower() == query or p[1].lower() == query
                     or str(p[0]) == query]
            if len(exact) == 1:
                self._apply_prop(exact[0])

    def on_search_enter(self, event=None):
        if self.filtered_props:
            self._apply_prop(self.filtered_props[0])

    def refresh_prop_list(self):
        self.prop_listbox.delete(0, tk.END)
        for pid, en, cn in self.filtered_props:
            self.prop_listbox.insert(tk.END, f"{pid:>3} | {cn}")

        self.match_label.config(text=f"匹配 {len(self.filtered_props)} / {len(self.props)} 个道具")

        if self.selected_prop:
            for i, p in enumerate(self.filtered_props):
                if p == self.selected_prop:
                    self.prop_listbox.selection_clear(0, tk.END)
                    self.prop_listbox.selection_set(i)
                    self.prop_listbox.see(i)
                    break

    def on_prop_select(self, event=None):
        sel = self.prop_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self.filtered_props):
            return
        self._apply_prop(self.filtered_props[idx])

    def on_prop_double_click(self, event=None):
        self.on_prop_select(event)

    def _apply_prop(self, prop):
        self.selected_prop = prop
        pid, en, cn = prop
        self.selected_label.config(text=f"已选：{pid} - {cn}", fg="#2E7D32")

    def clear_prop_search(self):
        self.prop_search.delete(0, tk.END)
        self.filtered_props = list(self.props)
        self.refresh_prop_list()

    def clear_prop_selection(self):
        self.selected_prop = None
        self.selected_label.config(text="未选择道具", fg="#9E9E9E")

    # ---------- 记录 ----------
    def record_combo(self):
        total = sum(self.quantities.values())
        if total != 8:
            messagebox.showwarning("提示", f"当前总数为 {total}，必须正好为 8。")
            return
        if self.selected_prop is None:
            messagebox.showwarning("提示", "请先选择一个道具。")
            return

        parts = []
        for _, name in self.ingredients:
            qty = self.quantities[name]
            if qty > 0:
                parts.append(f"{name}{qty}")
        recipe = "+".join(parts)

        pid, en, cn = self.selected_prop
        value = sum(self.quantities[n] * self.ingredient_values.get(n, 0)
                    for n in self.quantities)
        orb = "是" if self.orb_var.get() else "否"
        # 手动记录 = 已实际合成
        record = f"{recipe}|{pid}|{cn}|{value}|是|{orb}"

        existing_recipes = {parse_record(r)[0] for r in self.combos}
        if recipe in existing_recipes:
            self.recipe_label.config(text=f"🔁 该配方已记录过：{recipe}", fg="#F57C00")
            return

        save_combination(self.filename, record)
        self.combos.append(record)
        self.listbox.insert(tk.END, record)
        self.recipe_label.config(
            text=f"✅ 已记录：{recipe} → {pid} {cn}（价值 {value}）", fg="#2E7D32")

    def refresh_record_list(self):
        self.listbox.delete(0, tk.END)
        for r in self.combos:
            self.listbox.insert(tk.END, r)

    def clear_records(self):
        label = f"种子 {self.seed}" if self.seed else "无种子"
        if not messagebox.askyesno(
                "确认",
                f"确定要删除【{label}】的所有记录吗？\n文件：{self.filename}"):
            return
        if os.path.exists(self.filename):
            os.remove(self.filename)
        self.combos.clear()
        self.listbox.delete(0, tk.END)
        self._sync_cache.pop(self.filename, None)
        messagebox.showinfo("完成", "记录已清空。")


    # ---------- 模组数据同步 ----------
    def _poll_mod_data(self):
        try:
            # 存档还没生成时（例如刚装模组、还没合成过）定期重新查找
            if not self._sync_paths or \
                    not any(os.path.exists(p) for p in self._sync_paths):
                self._sync_retry += 1
                if self._sync_retry >= 4:
                    self._sync_retry = 0
                    self._sync_paths = find_mod_saves()
                    self._sync_sigs = {}

            found = [p for p in self._sync_paths if os.path.exists(p)]
            added = self.sync_from_mod() if found else 0
            if added:
                self._sync_total += added
            if found:
                names = "、".join(os.path.basename(p) for p in found)
                text = f"模组同步：已导入 {self._sync_total} 条 · {names}"
            else:
                text = "模组同步：未找到模组存档"
            if text != self._sync_status_text:
                self._sync_status_text = text
                self.sync_label.config(text=text)
        except Exception:
            pass
        self.root.after(1500, self._poll_mod_data)

    def sync_from_mod(self):
        """读取所有槽位的模组存档，返回本次新增（不含更新）的记录条数。"""
        added = 0
        for path in self._sync_paths:
            if not os.path.exists(path):
                continue
            # 模组会整份重写存档，因此用 (大小, 修改时间) 判断是否有变化，
            # 而不是只读增量——否则“否 -> 是”这类等长原地更新会被漏掉。
            st = os.stat(path)
            sig = (st.st_size, st.st_mtime_ns)
            if self._sync_sigs.get(path) == sig:
                continue
            self._sync_sigs[path] = sig
            added += self._sync_lines_from(path)
        return added

    def _sync_lines_from(self, path):
        """解析单个槽位存档，返回新增记录条数。

        模组每行格式：配方|道具ID|中文名|种子|价值|是否合成|十字圣球
        落盘格式：   配方|道具ID|中文名|价值|是否合成|十字圣球
        """
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()

        added = 0
        for line in content.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) < 2:
                continue
            recipe = parts[0]
            try:
                pid = int(parts[1])
            except ValueError:
                continue
            cn = self.prop_id_to_name.get(pid) or \
                (parts[2] if len(parts) > 2 else "")
            seed = normalize_seed(parts[3]) if len(parts) > 3 and parts[3] else ""
            value = parts[4] if len(parts) > 4 else \
                str(recipe_value(recipe, self.ingredient_values))
            crafted = parts[5] if len(parts) > 5 and parts[5] else "-"
            orb = parts[6] if len(parts) > 6 and parts[6] else "-"
            record = f"{recipe}|{pid}|{cn}|{value}|{crafted}|{orb}"
            if self._upsert_record(get_filename(seed), recipe, record) == "new":
                added += 1
        return added

    def _lines_for(self, filename):
        """记录文件在内存中的行列表（当前种子直接用 self.combos）"""
        if filename == self.filename:
            return self.combos
        if filename not in self._sync_cache:
            self._sync_cache[filename] = load_combinations(filename)
        return self._sync_cache[filename]

    def _upsert_record(self, filename, recipe, record):
        """按配方去重写入；已存在但内容不同则原地更新。

        返回 "new" / "updated" / None（无变化）。
        """
        lines = self._lines_for(filename)
        for i, line in enumerate(lines):
            if parse_record(line)[0] != recipe:
                continue

            # 同一个配方可能同时出现在多个槽位存档里，合并时：
            #   - “价值 / 十字圣球”不要用未知的 “-” 盖掉已有信息
            #   - “是否合成”只从否升级为是，不回退
            old = line.split("|")
            new = record.split("|")
            old += ["-"] * (len(new) - len(old))
            for k in (3, 5):
                if old[k] not in ("", "-") and new[k] in ("", "-"):
                    new[k] = old[k]
            if old[4] == "是":
                new[4] = "是"
            record = "|".join(new)

            if line == record:
                return None
            lines[i] = record
            with open(filename, "w", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")
            if filename == self.filename:
                self.refresh_record_list()
            return "updated"

        save_combination(filename, record)
        lines.append(record)
        if filename == self.filename:
            self.listbox.insert(tk.END, record)
        return "new"

    # ---------- 导入 / 导出 ----------
    def import_records(self):
        path = filedialog.askopenfilename(
            title="选择要导入的记录文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")])
        if not path:
            return
        existing = {parse_record(r)[0] for r in self.combos}
        added = 0
        for line in load_combinations(path):
            recipe, pid, cn = parse_record(line)
            if recipe and recipe not in existing:
                parts = line.split("|")
                if len(parts) < 6:
                    # 旧格式：补价值列，“是否合成 / 十字圣球”标为未知
                    parts = [recipe, pid, cn,
                             str(recipe_value(recipe, self.ingredient_values))]
                    parts += ["-"] * (6 - len(parts))
                record = "|".join(parts)
                save_combination(self.filename, record)
                self.combos.append(record)
                existing.add(recipe)
                added += 1
        self._sync_cache.pop(self.filename, None)
        self.refresh_record_list()
        messagebox.showinfo("导入完成", f"导入 {added} 条新记录。")

    def export_records(self):
        if not self.combos:
            messagebox.showinfo("提示", "当前没有可导出的记录。")
            return
        path = filedialog.asksaveasfilename(
            title="导出记录",
            defaultextension=".txt",
            initialfile=self.filename,
            filetypes=[("文本文件", "*.txt")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.combos) + "\n")
        messagebox.showinfo("导出完成",
                            f"已导出 {len(self.combos)} 条记录到：\n{path}")


if __name__ == "__main__":
    root = tk.Tk()

    # ---------- 动态选择跨平台字体 ----------
    def get_font(family_candidates, size, weight="normal"):
        available = set(tkfont.families())
        for family in family_candidates:
            if family in available:
                return tkfont.Font(family=family, size=size, weight=weight)
        return tkfont.Font(size=size, weight=weight)

    CN_FAMILIES = [
        "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC",
        "WenQuanYi Micro Hei", "WenQuanYi Zen Hei", "SimHei", "sans-serif",
    ]
    MONO_FAMILIES = [
        "Consolas", "DejaVu Sans Mono", "Noto Sans Mono",
        "Courier New", "monospace",
    ]

    FONT_NORMAL = get_font(CN_FAMILIES, 10)
    FONT_BOLD = get_font(CN_FAMILIES, 10, "bold")
    FONT_TITLE = get_font(CN_FAMILIES, 14, "bold")
    FONT_SMALL = get_font(CN_FAMILIES, 9)
    FONT_MONO = get_font(MONO_FAMILIES, 10)
    FONT_SEED = get_font(MONO_FAMILIES, 12)

    root.option_add("*Font", FONT_NORMAL)

    ingredients = load_ingredients()
    props = load_props()

    item_names_in_order = [name for _, name in ingredients]
    sprites, sprite_refs = load_sprites(
        SPRITE_SHEET, item_names_in_order, SPRITE_SIZE
    )

    app = ComboRecorder(root, ingredients, props, sprites)
    app._sprite_refs = sprite_refs   # 防止 PhotoImage 被 GC
    root.mainloop()
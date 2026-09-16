import os
import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook
import tkinter.font as tkfont

from seed import str2seed, is_valid_seed

INGREDIENTS_FILE = "合成宝袋组件.xlsx"
PROPS_FILE = "以撒的结合忏悔+_全道具信息表.xlsx"

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


# ---------- 主界面 ----------
class ComboRecorder:
    def __init__(self, root, ingredients, props):
        self.root = root
        self.root.title("合成宝袋配方记录器")
        self.root.geometry("700x820")
        self.root.configure(bg="#FAFAFA")
        self.root.resizable(True, True)

        self.ingredients = ingredients
        self.props = props
        self.filtered_props = list(props)
        self.quantities = {name: 0 for _, name in ingredients}
        self.selected_prop = None
        self.seed = ""
        self.filename = get_filename("")
        self.combos = load_combinations(self.filename)

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
                                     wraplength=640, justify="center")
        self.recipe_label.pack(pady=(4, 0))

        # ============ 物品按钮 ============
        btn_frame = tk.Frame(root, bg="#FAFAFA")
        btn_frame.grid(row=3, column=0, pady=2)

        self.buttons = {}
        COLS = 4
        for i, (order, name) in enumerate(ingredients):
            btn = tk.Button(btn_frame, text=f"{name}\n0",
                            width=13, height=2,
                            font=FONT_SMALL,
                            bg="#FFFFFF", activebackground="#E3F2FD",
                            relief="solid", bd=1, cursor="hand2",
                            command=lambda n=name: self.add_item(n))
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
        self.prop_listbox = tk.Listbox(prop_list_frame, height=6,
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
        bottom.grid(row=7, column=0, pady=8)

        for text, cmd in [
            ("记录配方", self.record_combo),
            ("清空当前", self.clear_current),
            ("删除记录文件", self.clear_records),
        ]:
            tk.Button(bottom, text=text, width=12,
                      font=FONT_NORMAL,
                      bg="#FFFFFF", activebackground="#E3F2FD",
                      relief="solid", bd=1, cursor="hand2",
                      command=cmd).pack(side="left", padx=5)

        # ============ 已记录列表 ============
        list_frame = tk.Frame(root, bg="#FAFAFA")
        list_frame.grid(row=8, column=0, padx=20, pady=(0, 12), sticky="nsew")
        root.grid_rowconfigure(8, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        tk.Label(list_frame, text="已记录：",
                 font=FONT_BOLD,
                 bg="#FAFAFA", fg="#424242").grid(row=0, column=0, sticky="w")

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
            btn.config(text=f"{name}\n{self.quantities[name]}")

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
        record = f"{recipe}|{pid}|{cn}"

        existing_recipes = {parse_record(r)[0] for r in self.combos}
        if recipe in existing_recipes:
            self.recipe_label.config(text=f"🔁 该配方已记录过：{recipe}", fg="#F57C00")
            return

        save_combination(self.filename, record)
        self.combos.append(record)
        self.listbox.insert(tk.END, record)
        self.recipe_label.config(text=f"✅ 已记录：{recipe} → {pid} {cn}", fg="#2E7D32")

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
        messagebox.showinfo("完成", "记录已清空。")


if __name__ == "__main__":
    root = tk.Tk()

    # ---------- 动态选择跨平台字体 ----------
    def get_font(family_candidates, size, weight="normal"):
        """根据系统可用字体返回 tkfont.Font 对象"""
        available = set(tkfont.families())
        for family in family_candidates:
            if family in available:
                return tkfont.Font(family=family, size=size, weight=weight)
        # 如果都没找到，返回默认字体
        return tkfont.Font(size=size, weight=weight)

    # 中文字体候选（按优先级）
    CN_FAMILIES = [
        "Microsoft YaHei",      # Windows
        "PingFang SC",          # macOS
        "Noto Sans CJK SC",     # Ubuntu (fonts-noto-cjk)
        "WenQuanYi Micro Hei",  # Ubuntu (fonts-wqy-microhei)
        "WenQuanYi Zen Hei",    # Ubuntu (fonts-wqy-zenhei)
        "SimHei",               # Windows 备选
        "sans-serif",           # 保底
    ]

    # 等宽字体候选（按优先级）
    MONO_FAMILIES = [
        "Consolas",             # Windows
        "DejaVu Sans Mono",     # Ubuntu 常见
        "Noto Sans Mono",       # Ubuntu 常见
        "Courier New",          # 跨平台备选
        "monospace",            # 保底
    ]

    # 生成全局字体对象
    FONT_NORMAL = get_font(CN_FAMILIES, 10)
    FONT_BOLD = get_font(CN_FAMILIES, 10, "bold")
    FONT_TITLE = get_font(CN_FAMILIES, 14, "bold")
    FONT_SMALL = get_font(CN_FAMILIES, 9)
    FONT_MONO = get_font(MONO_FAMILIES, 10)
    FONT_SEED = get_font(MONO_FAMILIES, 12)

    # 设置全局默认字体
    root.option_add("*Font", FONT_NORMAL)

    ingredients = load_ingredients()
    props = load_props()
    app = ComboRecorder(root, ingredients, props)
    root.mainloop()
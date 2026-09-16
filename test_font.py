import tkinter as tk
import tkinter.font as tkfont

root = tk.Tk()
available = set(tkfont.families())

candidates = [
    "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC",
    "WenQuanYi Micro Hei", "WenQuanYi Zen Hei", "SimHei",
    "DejaVu Sans Mono", "Noto Sans Mono", "Consolas"
]

for f in candidates:
    print(f"{f}: {'✅ 可用' if f in available else '❌ 不可用'}")

root.destroy()
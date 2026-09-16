import tkinter as tk
import tkinter.font as tkfont
import sys

print("Python:", sys.executable)
print("Tk 版本:", tk.TkVersion)

root = tk.Tk()
fams = tkfont.families()
print("字体总数:", len(fams))
print("前 30 个:", fams[:30])
print("默认字体:", tkfont.nametofont("TkDefaultFont").actual())
root.destroy()
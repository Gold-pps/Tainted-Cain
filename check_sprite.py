from PIL import Image

img = Image.open("Crafting_ui_sprite.png")
w, h = img.size
print(f"图片尺寸: {w} x {h}")

# 试试常见的网格
for cols, rows in [(8, 3), (8, 4), (7, 4), (4, 7), (6, 5)]:
    cw, ch = w / cols, h / rows
    print(f"{cols}列 x {rows}行 -> 每格 {cw:.1f} x {ch:.1f}")
from PIL import Image, ImageDraw

img = Image.open("Crafting_ui_sprite.png").convert("RGBA")
w, h = img.size

SCALE = 4
big = img.resize((w * SCALE, h * SCALE), Image.NEAREST)

draw = ImageDraw.Draw(big)

# 每 16 像素画一条网格线（对应原图 16px）
for x in range(0, w + 1, 16):
    draw.line([(x * SCALE, 0), (x * SCALE, h * SCALE)], fill=(255, 0, 0, 128), width=1)
    draw.text((x * SCALE + 2, 2), str(x), fill=(255, 0, 0))

for y in range(0, h + 1, 16):
    draw.line([(0, y * SCALE), (w * SCALE, y * SCALE)], fill=(0, 128, 255, 128), width=1)
    draw.text((2, y * SCALE + 2), str(y), fill=(0, 128, 255))

big.save("grid_debug.png")
print("已生成 grid_debug.png，尺寸", big.size)
print(f"原图 {w}x{h}，网格每格 16x16")
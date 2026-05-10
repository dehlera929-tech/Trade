from PIL import Image, ImageDraw, ImageFont

W, H = 1080, 1350
img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
font_pov = ImageFont.truetype(FONT, 90)
font_main = ImageFont.truetype(FONT, 52)
font_pace = ImageFont.truetype(FONT, 56)
font_stats = ImageFont.truetype(FONT, 52)

pad = 50
box_top = 70
box_h = 470
draw.rounded_rectangle(
    [(pad, box_top), (W - pad, box_top + box_h)],
    radius=36,
    fill=(0, 0, 0, 195),
)

def center(text, font, y, fill=(255, 255, 255, 255)):
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    draw.text(((W - tw) / 2, y), text, font=font, fill=fill)

center("POV:", font_pov, box_top + 30, fill=(255, 200, 0, 255))
center("Du gehst 3x die Woche ins Gym", font_main, box_top + 160)
center("und bist trotzdem unsportlich.", font_main, box_top + 230)
center("5:01er Pace", font_pace, box_top + 340, fill=(255, 200, 0, 255))
center("auf 10 km", font_pace, box_top + 405, fill=(255, 200, 0, 255))

bottom_box_h = 160
bb_top = H - bottom_box_h - 70
draw.rounded_rectangle(
    [(pad, bb_top), (W - pad, bb_top + bottom_box_h)],
    radius=36,
    fill=(0, 0, 0, 195),
)
center("10,01 km  ·  50:15  ·  5:01 /km", font_stats, bb_top + 50)

img.save("/home/user/Trade/overlay.png")
print("OK overlay.png")

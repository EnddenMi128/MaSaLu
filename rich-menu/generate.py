"""Generate a LINE calendar image and tap areas for 14 consecutive dates."""
import argparse
from datetime import date, timedelta
import calendar
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

WIDTH, HEIGHT = 2500, 843
LIFF = "https://liff.line.me/2010020470-VBPZRUQ9/"
BG, INK, ACCENT = "#16100D", "#FFF3E5", "#FF873D"


def generate(start, font_path, font_index):
    dates = [start + timedelta(days=offset) for offset in range(14)]
    end = dates[-1]
    output = Path(__file__).parent
    background = Image.open(output / "masalu-background.jpg").convert("RGB")
    image = Image.blend(background, Image.new("RGB", (WIDTH, HEIGHT), BG), 0.57)
    draw = ImageDraw.Draw(image)
    fonts = {size: ImageFont.truetype(font_path, size, index=font_index) for size in (44, 54, 66)}
    heading = ImageFont.truetype("DejaVuSerif-Bold.ttf", 112)
    areas = []

    def area(box, label, uri):
        x, y, right, bottom = box
        areas.append({"bounds": {"x": x, "y": y, "width": right-x, "height": bottom-y},
                      "action": {"type": "uri", "label": label, "uri": uri}})

    month = calendar.month_name[start.month]
    if end.month != start.month:
        month += " / " + calendar.month_abbr[end.month]
    year = str(start.year) if start.year == end.year else f"{start.year}–{end.year}"
    draw.text((35, 84), month, font=heading, fill=ACCENT, anchor="lm")
    draw.text((2465, 84), year, font=heading, fill=ACCENT, anchor="rm")

    for column, day in enumerate(dates[:7]):
        x = 25 + column * 352
        draw.rounded_rectangle((x, 164, x + 338, 252), radius=8, outline=INK, width=2)
        draw.text((x + 169, 208), calendar.day_abbr[day.weekday()].upper(),
                  font=fonts[54], fill=INK, anchor="mm")

    for offset, day in enumerate(dates):
        row, column = divmod(offset, 7)
        x, y = 25 + column * 352, 264 + row * 233
        box = (x, y, x + 338, y + 221)
        color = ACCENT if offset == 0 else INK
        draw.rounded_rectangle(box, radius=8, outline=color, width=4 if offset == 0 else 2)
        draw.text((x + 26, y + 46), f"{day.month}/{day.day}", font=fonts[66], fill=color, anchor="lm",
                  stroke_width=2, stroke_fill=BG)
        area(box, f"{day.month} 月 {day.day} 日", f"{LIFF}?date={day.isoformat()}&from=richmenu")

    box = (25, 733, 2475, 824)
    draw.rounded_rectangle(box, radius=12, fill=ACCENT)
    draw.text((1250, 778), "開啟完整月曆", font=fonts[44], fill=BG, anchor="mm")
    area(box, "開啟完整月曆", LIFF)
    image_name = f"calendar-{start.isoformat()}-masalu.jpg"
    image.save(output / image_name, quality=88, optimize=True)
    manifest = {"alias": f"masalu-art-{start.isoformat()}", "image": image_name,
                "start": start.isoformat(), "end": end.isoformat(),
                "menu": {"size": {"width": WIDTH, "height": HEIGHT}, "selected": True,
                         "name": f"MaSaLu {start.isoformat()} – {end.isoformat()}",
                         "chatBarText": "排班月曆", "areas": areas}}
    (output / "menu.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("start", type=date.fromisoformat, help="YYYY-MM-DD")
    parser.add_argument("--font", default="/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc")
    parser.add_argument("--font-index", default=3, type=int)
    args = parser.parse_args()
    generate(args.start, args.font, args.font_index)

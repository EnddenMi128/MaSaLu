"""Generate a LINE calendar image and tap areas for 14 consecutive dates."""
import argparse
from datetime import date, timedelta
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

WIDTH, HEIGHT = 2500, 1250
LIFF = "https://liff.line.me/2010020470-VBPZRUQ9/"
BG, CARD, INK, MUTED, ACCENT = "#16100D", "#2C1E18", "#FFF3E5", "#D8B9A3", "#FF873D"


def generate(start, font_path, font_index):
    dates = [start + timedelta(days=offset) for offset in range(14)]
    end = dates[-1]
    output = Path(__file__).parent
    image = Image.new("RGB", (WIDTH, HEIGHT), BG)
    draw = ImageDraw.Draw(image)
    fonts = {size: ImageFont.truetype(font_path, size, index=font_index) for size in (56, 70, 95, 110)}
    areas = []

    def area(box, label, uri):
        x, y, right, bottom = box
        areas.append({"bounds": {"x": x, "y": y, "width": right-x, "height": bottom-y},
                      "action": {"type": "uri", "label": label, "uri": uri}})

    draw.text((88, 105), "MaSaLu  排班月曆", font=fonts[95], fill=INK, anchor="lm")
    draw.text((2412, 105), f"{start:%m/%d} – {end:%m/%d}", font=fonts[70], fill=ACCENT, anchor="rm")
    draw.line((88, 191, 2412, 191), fill=ACCENT, width=5)
    for offset, day in enumerate(dates):
        row, column = divmod(offset, 7)
        x, y = 88 + column * 332, 244 + row * 377
        box = (x + 7, y, x + 325, y + 345)
        color = ACCENT if day.weekday() >= 5 or offset == 0 else INK
        draw.rounded_rectangle(box, radius=26, fill=CARD, outline=color if offset == 0 else "#71503D", width=5)
        draw.text((x + 166, y + 77), "週" + "一二三四五六日"[day.weekday()], font=fonts[56], fill=MUTED, anchor="mm")
        draw.text((x + 166, y + 199), f"{day.month}/{day.day}", font=fonts[95], fill=color, anchor="mm")
        area(box, f"{day.month} 月 {day.day} 日", f"{LIFF}?date={day.isoformat()}&from=richmenu")

    box = (88, 1032, 2412, 1165)
    draw.rounded_rectangle(box, radius=28, fill=ACCENT)
    draw.text((1250, 1098), "開啟完整月曆", font=fonts[56], fill=BG, anchor="mm")
    area(box, "開啟完整月曆", LIFF)
    image_name = f"calendar-{start.isoformat()}.png"
    image.save(output / image_name, optimize=True)
    manifest = {"alias": f"masalu-{start.isoformat()}", "image": image_name,
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

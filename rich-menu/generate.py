"""Generate a LINE calendar image and tap areas for 14 consecutive dates."""
import argparse
import base64
from datetime import date, timedelta
import calendar
import gzip
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps

WIDTH, HEIGHT = 2500, 1686
PERSONAL_WIDTH, PERSONAL_HEIGHT = 1200, 809
LIFF = "https://liff.line.me/2010020470-VBPZRUQ9/"
BG, INK, ACCENT = "#16100D", "#FFF3E5", "#FF873D"


def generate(today, font_path, font_index):
    start = today - timedelta(days=(today.weekday() + 1) % 7)
    dates = [start + timedelta(days=offset) for offset in range(14)]
    end = dates[-1]
    output = Path(__file__).parent
    background = Image.open(output / "masalu-background.jpg").convert("RGB")
    background = ImageOps.fit(background, (WIDTH, HEIGHT), Image.Resampling.LANCZOS)
    image = Image.blend(background, Image.new("RGB", (WIDTH, HEIGHT), BG), 0.57)
    draw = ImageDraw.Draw(image)
    fonts = {size: ImageFont.truetype(font_path, size, index=font_index) for size in (56, 68, 76, 82, 128)}
    heading = ImageFont.truetype("DejaVuSerif-Bold.ttf", 124)
    areas = []

    def area(box, label, uri):
        x, y, right, bottom = box
        areas.append({"bounds": {"x": x, "y": y, "width": right-x, "height": bottom-y},
                      "action": {"type": "uri", "label": label, "uri": uri}})

    month = calendar.month_name[start.month]
    if end.month != start.month:
        month += " / " + calendar.month_abbr[end.month]
    year = str(start.year) if start.year == end.year else f"{start.year}–{end.year}"
    draw.text((35, 105), month, font=heading, fill=ACCENT, anchor="lm")
    draw.text((2040, 105), year, font=heading, fill=ACCENT, anchor="rm")

    for column, day in enumerate(dates[:7]):
        x = 25 + column * 352
        draw.rounded_rectangle((x, 197, x + 338, 322), radius=8, outline=INK, width=2)
        draw.text((x + 169, 259), calendar.day_abbr[day.weekday()].upper(),
                  font=fonts[82], fill=INK, anchor="mm")

    for offset, day in enumerate(dates):
        row, column = divmod(offset, 7)
        x, y = 25 + column * 352, 338 + row * 581
        box = (x, y, x + 338, y + 563)
        color = ACCENT if day == start else INK
        draw.rounded_rectangle(box, radius=8, outline=color, width=4 if day == start else 2)
        draw.text((x + 24, y + 75), str(day.day), font=fonts[128], fill=color, anchor="lm",
                  stroke_width=2, stroke_fill=BG)
        area(box, f"{day.month} 月 {day.day} 日", f"{LIFF}?date={day.isoformat()}&from=richmenu")

    footer = ["開啟完整月曆", "聯絡團員", "請假審核", "回覆確認", "缺班查詢"]
    for index, label in enumerate(footer):
        x = 25 + index * 492
        box = (x, 1500, x + 482, 1660)
        draw.rounded_rectangle(box, radius=12, fill=ACCENT)
        draw.text((x + 241, 1580), label, font=fonts[68], fill=BG, anchor="mm")
        if index == 0:
            area(box, label, LIFF)
        else:
            areas.append({
                "bounds": {"x": x, "y": 1500, "width": 482, "height": 160},
                "action": {"type": "message", "label": label, "text": label},
            })
    personal_button = (2070, 42, 2475, 168)
    draw.rounded_rectangle(personal_button, radius=16, fill=ACCENT)
    draw.text((2272, 105), "更新/個人班表", font=fonts[56], fill=BG, anchor="mm")
    areas.append({
        "bounds": {"x": 2070, "y": 42, "width": 405, "height": 126},
        "action": {"type": "message", "label": "更新/個人班表", "text": "個人班表"},
    })
    image_name = f"calendar-{start.isoformat()}-tall-v6.jpg"
    image.save(output / image_name, quality=88, optimize=True)
    personal = image.resize((PERSONAL_WIDTH, PERSONAL_HEIGHT), Image.Resampling.LANCZOS).quantize(colors=236)
    palette = personal.getpalette()[:236 * 3] + [0] * (16 * 3)
    palette += [16, 12, 10, 255, 243, 229, 255, 135, 61, 255, 210, 145]
    glyphs = {}
    font = ImageFont.truetype(font_path, 23, index=font_index)
    for label in ("已確認", "待回覆"):
        mask = Image.new("L", (72, 32))
        ImageDraw.Draw(mask).text((0, 0), label, font=font, fill=255)
        glyphs[label] = base64.b64encode(bytes(value >= 96 for value in mask.tobytes())).decode()
    personal_name = f"personal-grid-tall-v6-{start.isoformat()}.json"
    (output / personal_name).write_text(json.dumps({
        "width": PERSONAL_WIDTH, "height": PERSONAL_HEIGHT,
        "pixels": base64.b64encode(gzip.compress(personal.tobytes(), mtime=0)).decode(),
        "palette": palette, "glyphs": glyphs
    }, ensure_ascii=False), encoding="utf-8")
    manifest = {"alias": f"masalu-tall-v6-{start.isoformat()}", "image": image_name,
                "personal": personal_name,
                "start": start.isoformat(), "end": end.isoformat(),
                "menu": {"size": {"width": WIDTH, "height": HEIGHT}, "selected": True,
                         "name": f"MaSaLu {start.isoformat()} – {end.isoformat()}",
                         "chatBarText": "排班月曆", "areas": areas}}
    (output / "menu.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def generate_roster_glyphs(font_path, font_index):
    font = ImageFont.truetype(font_path, 16, index=font_index)
    glyphs = bytearray()
    for code in [*range(32, 127), *range(0x3400, 0xA000)]:
        mask = Image.new("1", (16, 18))
        ImageDraw.Draw(mask).text((8, 9), chr(code), font=font, fill=1, anchor="mm")
        glyphs.extend(mask.tobytes())
    path = Path(__file__).parent / "roster-glyphs.gz"
    path.write_bytes(gzip.compress(glyphs, mtime=0))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("today", type=date.fromisoformat, help="Any date in the first week (YYYY-MM-DD)")
    parser.add_argument("--font", default="/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc")
    parser.add_argument("--font-index", default=3, type=int)
    parser.add_argument("--roster-font", action="store_true", help="Also regenerate the roster glyph atlas")
    args = parser.parse_args()
    generate(args.today, args.font, args.font_index)
    if args.roster_font:
        generate_roster_glyphs(args.font, args.font_index)

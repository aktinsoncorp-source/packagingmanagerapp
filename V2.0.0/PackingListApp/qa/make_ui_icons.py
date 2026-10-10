"""Generate the small, matching UI icon set; no font glyph dependencies."""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ASSETS = Path(__file__).resolve().parents[1] / "src" / "assets" / "ui"
SCALE = 4
INK, BLUE, TEAL, GRAY = "#24292c", "#42484b", "#006e65", "#6b767c"


def render(name, size):
    image = Image.new("RGBA", (size * SCALE, size * SCALE))
    draw = ImageDraw.Draw(image)

    def box(values):
        return tuple(round(v * SCALE) for v in values)

    def stroke(points, color, width):
        draw.line([box(p) for p in points], fill=color, width=round(width * SCALE), joint="curve")
        radius = width / 2
        for x, y in (points[0], points[-1]):
            draw.ellipse(box((x-radius, y-radius, x+radius, y+radius)), fill=color)

    if name.startswith("check"):
        state = name.removeprefix("check-")
        active = state in ("on", "mixed")
        color = "#bfc3c7" if active else GRAY if state == "disabled" else INK
        fill = BLUE if active else "#fcfcf9"
        draw.rounded_rectangle(box((2, 2, 22, 22)), radius=5*SCALE, fill=fill, outline=color, width=2*SCALE)
        if state == "on":
            stroke([(6.5, 12), (10.2, 15.7), (17.6, 8.2)], "white", 2.5)
        elif state in ("mixed", "disabled"):
            stroke([(7, 12), (17, 12)], "white" if active else GRAY, 2.5)
    else:
        opened = "open" in name
        color = GRAY if "disabled" in name else TEAL if opened else INK
        if "header" in name:
            color = "#48d1b6" if opened else "#fcfcf9"
        # Both bodies share a baseline; the open shackle has a clear gap.
        if opened:
            draw.arc(box((12, 3, 30, 21)), 180, 360, fill=color, width=round(2.5*SCALE))
            stroke([(13.2, 12), (13.2, 17)], color, 2.5)
            stroke([(28.8, 12), (28.8, 12.5)], color, 2.5)
        else:
            draw.arc(box((9, 3, 23, 17)), 180, 360, fill=color, width=round(2.5*SCALE))
            stroke([(10.2, 10), (10.2, 17)], color, 2.5)
            stroke([(21.8, 10), (21.8, 17)], color, 2.5)
        draw.rounded_rectangle(box((5, 15, 27, 29)), radius=4*SCALE,
                               fill=BLUE if "header" in name else "#fcfcf9",
                               outline=color, width=round(2.5*SCALE))
    return image.resize((size, size), Image.Resampling.LANCZOS)


def main():
    ASSETS.mkdir(parents=True, exist_ok=True)
    names = ("check-off", "check-on", "check-mixed", "check-disabled",
             "lock-closed", "lock-open", "lock-closed-disabled", "lock-open-disabled")
    preview = Image.new("RGB", (760, 200), "#f1f2ef")
    draw = ImageDraw.Draw(preview)
    font = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", 15)
    for index, name in enumerate(names):
        icon = render(name, 24 if name.startswith("check") else 32)
        icon.save(ASSETS / f"{name}.png")
        if name in ("lock-closed", "lock-open"):
            render(name+"-header", 32).save(ASSETS / f"{name}-header.png")
        x, y = 15+(index % 4)*190, 16+(index // 4)*94
        preview.paste(icon, (x, y), icon)
        draw.text((x, y+40), name.replace("check-", "").replace("lock-", ""), fill=INK, font=font)
    output = Path(__file__).resolve().parent / "ui-icons-preview.png"
    preview.save(output)
    print(output)


if __name__ == "__main__":
    main()

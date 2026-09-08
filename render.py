"""
render.py
---------
Draws flags onto a rendered page image. Pure PIL, no Streamlit.
"""

import io
from PIL import Image, ImageDraw

SOURCE_COLOR = {"rules": (220, 38, 38), "dictionary": (217, 119, 6), "llm": (124, 58, 237)}
SOURCE_LABEL = {"rules": "glossary", "dictionary": "dictionary", "llm": "AI"}


def render_page_image(page, dpi=150):
    pix = page.get_pixmap(dpi=dpi)
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def render_flagged_page(page, spell_flags, spacing_flags, dpi=150):
    scale = dpi / 72
    img = render_page_image(page, dpi=dpi)
    draw = ImageDraw.Draw(img)

    n = 0
    for f in spell_flags:
        n += 1
        if not f.get("rect"):
            continue
        r = f["rect"]
        x0, y0, x1, y1 = r.x0 * scale, r.y0 * scale, r.x1 * scale, r.y1 * scale
        pad = 15
        color = SOURCE_COLOR[f["source"]]
        draw.ellipse([x0 - pad, y0 - pad, x1 + pad, y1 + pad], outline=color, width=5)
        label = f"#{n} '{f['word']}' -> '{f['suggestion']}'? [{SOURCE_LABEL[f['source']]}]"
        tw = len(label) * 8
        ly0, ly1 = (y0 + y1) / 2 - 12, (y0 + y1) / 2 + 12
        draw.rectangle([x1 + pad, ly0, x1 + pad + tw, ly1], fill=color)
        draw.text((x1 + pad + 3, ly0 + 4), label, fill=(255, 255, 255))

    for f in spacing_flags:
        n += 1
        if not f.get("rect"):
            continue
        r = f["rect"]
        x0, y0, x1, y1 = r.x0 * scale, r.y0 * scale, r.x1 * scale, r.y1 * scale
        pad = 10
        color = SOURCE_COLOR[f["source"]]
        draw.rectangle([x0 - pad, y0 - pad, x1 + pad, y1 + pad], outline=color, width=4)
        label = f"#{n} {f['detail']} [{SOURCE_LABEL[f['source']]}]"
        tw = len(label) * 8
        ly0, ly1 = y1 + pad, y1 + pad + 20
        draw.rectangle([x0 - pad, ly0, x0 - pad + tw, ly1], fill=color)
        draw.text((x0 - pad + 3, ly0 + 3), label, fill=(255, 255, 255))

    return img
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw


folder = Path(sys.argv[1])
prefix = sys.argv[2]
files = sorted(folder.glob(f"{prefix}-*.png"))
for batch, start in enumerate(range(0, len(files), 4), 1):
    images = [Image.open(p).convert("RGB") for p in files[start:start + 4]]
    thumb_w = 900
    thumbs = []
    for p, image in zip(files[start:start + 4], images):
        h = round(image.height * thumb_w / image.width)
        resized = image.resize((thumb_w, h), Image.Resampling.LANCZOS)
        canvas = Image.new("RGB", (thumb_w, h + 44), "white")
        canvas.paste(resized, (0, 44))
        ImageDraw.Draw(canvas).text((12, 12), p.name, fill="black")
        thumbs.append(canvas)
    cell_h = max(i.height for i in thumbs)
    sheet = Image.new("RGB", (thumb_w * 2, cell_h * 2), "#d9d9d9")
    for i, image in enumerate(thumbs):
        sheet.paste(image, ((i % 2) * thumb_w, (i // 2) * cell_h))
    out = folder / f"{prefix}-contact-{batch:02d}.png"
    sheet.save(out, quality=92)
    print(out)

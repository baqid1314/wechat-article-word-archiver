from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageDraw


def main() -> None:
    parser = argparse.ArgumentParser(description="Create 2x2 contact sheets from rendered DOCX pages.")
    parser.add_argument("--render-dir", required=True)
    parser.add_argument("--output-dir", required=True)
    args = parser.parse_args()
    render_dir = Path(args.render_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    total_pages = total_sheets = 0
    for folder in sorted(path for path in render_dir.iterdir() if path.is_dir()):
        pages = sorted(folder.glob("page-*.png"))
        if not pages:
            raise RuntimeError(f"no rendered pages in {folder}")
        ink: list[float] = []
        for page in pages:
            with Image.open(page) as source:
                gray = source.convert("L").resize((100, 130))
                ink.append(sum(gray.histogram()[:245]) / 13000)
        print(f"{folder.name}\tpages={len(pages)}\tink_min={min(ink):.4f}\tink_max={max(ink):.4f}")
        for start in range(0, len(pages), 4):
            group = pages[start:start + 4]
            images = [Image.open(page).convert("RGB") for page in group]
            width = max(image.width for image in images)
            height = max(image.height for image in images)
            label_height = 28
            sheet = Image.new("RGB", (width * 2, (height + label_height) * 2), "#B8B8B8")
            draw = ImageDraw.Draw(sheet)
            for index, (page, image) in enumerate(zip(group, images)):
                x = (index % 2) * width
                y = (index // 2) * (height + label_height)
                sheet.paste(image, (x, y + label_height))
                draw.rectangle((x, y, x + width, y + label_height), fill="white")
                draw.text((x + 10, y + 7), page.stem, fill="black")
                image.close()
            destination = output_dir / f"{folder.name}__{start + 1:03d}-{start + len(group):03d}.jpg"
            sheet.save(destination, "JPEG", quality=88, optimize=True)
            sheet.close()
            total_sheets += 1
        total_pages += len(pages)
    print(f"TOTAL\tpages={total_pages}\tsheets={total_sheets}")


if __name__ == "__main__":
    main()

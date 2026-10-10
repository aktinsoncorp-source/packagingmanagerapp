from pathlib import Path
import pypdfium2 as pdfium
from PIL import Image, ImageOps, ImageDraw
import sys

root = Path(sys.argv[1]) if len(sys.argv)>1 else Path(__file__).parent / 'reference'
for file in root.glob('*.pdf'):
    pdf = pdfium.PdfDocument(file)
    thumbs = []
    for n, page in enumerate(pdf, 1):
        img = page.render(scale=1.6).to_pil().convert('RGB')
        out = root / file.stem
        out.mkdir(exist_ok=True)
        img.save(out / f'page-{n}.png')
        thumb = ImageOps.contain(img, (475, 672))
        tile = Image.new('RGB', (495, 710), '#d9dfe5')
        tile.paste(thumb, (10, 28))
        ImageDraw.Draw(tile).text((10, 8), f'{file.stem} page {n}', fill='black')
        thumbs.append(tile)
    cols = min(3, len(thumbs))
    sheet = Image.new('RGB', (495 * cols, 710 * ((len(thumbs) + cols - 1) // cols)), 'white')
    for i, thumb in enumerate(thumbs):
        sheet.paste(thumb, ((i % cols) * 495, (i // cols) * 710))
    sheet.save(root / f'{file.stem}-overview.png')
    print(file.name, len(pdf))

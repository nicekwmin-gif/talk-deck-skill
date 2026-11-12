"""Thumbnail grid of rendered slides (slide-XX.png from render_pptx.ps1) — one image to check flow or send to a phone.

Usage: python overview_grid.py png_dir out.jpg [--cols 4] [--width 400]
"""
import argparse, glob, os
from PIL import Image, ImageDraw, ImageFont

ap = argparse.ArgumentParser()
ap.add_argument('png_dir'); ap.add_argument('out')
ap.add_argument('--cols', type=int, default=4); ap.add_argument('--width', type=int, default=400)
a = ap.parse_args()
files = sorted(glob.glob(os.path.join(a.png_dir, 'slide-*.png')))
assert files, 'no slide-*.png found'
w = a.width; h = int(w * 9 / 16); pad = 14; lab = 22
rows = (len(files) + a.cols - 1) // a.cols
grid = Image.new('RGB', (a.cols * (w + pad) + pad, rows * (h + pad + lab) + pad), 'white')
d = ImageDraw.Draw(grid)
try: font = ImageFont.truetype('malgun.ttf', 15)
except OSError: font = ImageFont.load_default()
for k, f in enumerate(files):
    im = Image.open(f).convert('RGB').resize((w, h), Image.LANCZOS)
    x = pad + (k % a.cols) * (w + pad); y = pad + (k // a.cols) * (h + pad + lab)
    d.text((x, y), os.path.splitext(os.path.basename(f))[0].replace('slide-', ''), fill='#374151', font=font)
    grid.paste(im, (x, y + lab)); d.rectangle([x, y + lab, x + w - 1, y + lab + h - 1], outline='#D1D5DB')
grid.save(a.out, quality=88)
print('saved', a.out, len(files), 'slides')

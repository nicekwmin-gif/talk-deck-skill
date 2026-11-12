"""Save only some slides of a deck as a new pptx (for "수정한 슬라이드만 보내 줘").

Usage: python export_slides.py deck.pptx out.pptx 3,8,13-18
Slide numbers are 1-based; ranges allowed. The input file is not modified.
"""
import os, sys
from pptx import Presentation
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pptx_kit as K

src, out, spec = sys.argv[1:4]
keep = set()
for part in spec.split(','):
    a, _, b = part.partition('-')
    keep.update(range(int(a), int(b or a) + 1))
prs = Presentation(src)
n = len(prs.slides)
K.keep_slides(prs, keep)
prs.save(out)
print(f'saved {out}: kept {sorted(keep)} of {n}')

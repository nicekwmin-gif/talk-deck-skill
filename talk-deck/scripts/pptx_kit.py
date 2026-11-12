"""pptx_kit — python-pptx helpers for talk decks (16:9, 13.333 x 7.5 in).

Usage in a builder script:
    import sys; sys.path.insert(0, r'<skill>/scripts')
    import pptx_kit as K
    prs = K.new_deck()                     # or Presentation(<copy of the user's pptx>)
    K.THEME.update(accent='2E7D4F')        # optional: change palette / font before drawing
    s = K.content_slide(prs, 'Finding 1 · Flavonoids', 'Cold acclimation raises flavonoids', source='Data: ...')
    K.text(s, 0.6, 1.6, 6, 1, [['Plain ', ('italic', {'italic': True})]], size=18)
    K.notes(prs, s, '한국어 대본 ...')
    K.renumber(prs); prs.save(out)

Conventions:
  - shape names 'Section tag', 'Title', 'Source', 'Slide number' are fixed so later scripts can find them.
  - text(): paras = str | [paragraph]; paragraph = str | [run]; run = str | (str, {size,bold,italic,color,sup,sub,font}).
  - every helper accepts a group shape as `slide`, so a row can be built inside s.shapes.add_group_shape().
"""
import copy, os
from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

E = 914400  # EMU per inch
THEME = dict(font='Calibri', navy='14213D', accent='2E7D4F', ink='1F2937', grey='6B7280',
             grey_l='F3F4F6', line='D1D5DB', berry='8E2C48', cache=None)
rgb = RGBColor.from_string


# ------------------------------------------------------------------ deck
def new_deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    return prs


def blank_layout(prs):
    for l in prs.slide_layouts:
        if l.name.lower() == 'blank': return l
    return prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]


# ------------------------------------------------------------------ primitives
def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None, lw=1.0, dash=None, name=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = rgb(fill)
    else: s.fill.background()
    if line:
        s.line.color.rgb = rgb(line); s.line.width = Pt(lw)
        if dash: s.line.dash_style = dash
    else: s.line.fill.background()
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE: s.adjustments[0] = radius
    s.shadow.inherit = False
    if name: s.name = name
    tf = s.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return s


def _east_asian(run, typeface):
    """Also set the East Asian font, so Korean/Japanese/Chinese text uses the deck font instead of a fallback."""
    rPr = run._r.get_or_add_rPr()
    for el in rPr.findall(qn('a:ea')): rPr.remove(el)
    ea = etree.SubElement(rPr, qn('a:ea')); ea.set('typeface', typeface)
    latin = rPr.find(qn('a:latin'))
    if latin is not None: latin.addnext(ea)
    if any('가' <= ch <= '힣' for ch in run.text):   # Hangul: with latinLnBrk=0 this keeps words whole
        rPr.set('lang', 'ko-KR'); rPr.set('altLang', 'en-US')


def text(slide, x, y, w, h, paras, size=16, color=None, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         italic=False, space_after=0, line_spacing=None, margin=0, shape=None, name=None):
    """Add a text box (or fill `shape`, e.g. a rect, replacing its text)."""
    color = color or THEME['ink']
    tb = shape if shape is not None else slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(margin); tf.margin_top = tf.margin_bottom = Inches(margin * 0.6)
    tf.vertical_anchor = anchor
    if shape is not None:
        for p_ in list(tf.paragraphs)[1:]: p_._p.getparent().remove(p_._p)
        for r_ in list(tf.paragraphs[0].runs): r_._r.getparent().remove(r_._r)
    if isinstance(paras, str): paras = [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p._p.get_or_add_pPr().set('latinLnBrk', '0')   # Korean: wrap at spaces, never inside a word
        if space_after: p.space_after = Pt(space_after)
        if line_spacing: p.line_spacing = line_spacing
        for r in (para if isinstance(para, list) else [para]):
            t, o = (r, {}) if isinstance(r, str) else r
            run = p.add_run(); run.text = t
            f = run.font; f.name = o.get('font', THEME['font']); f.size = Pt(o.get('size', size))
            _east_asian(run, o.get('font', THEME['font']))
            f.bold = o.get('bold', bold); f.italic = o.get('italic', italic); f.color.rgb = rgb(o.get('color', color))
            if o.get('sub'): run.font._element.set('baseline', '-25000')
            if o.get('sup'): run.font._element.set('baseline', '30000')
    if name: tb.name = name
    return tb


def bullets(tb, color, char='•', indent_in=0.32):
    """Turn every paragraph of a text box into a hanging-indent bullet (real bullets, not typed '•')."""
    for p in tb.text_frame.paragraphs:
        pPr = p._p.get_or_add_pPr()
        pPr.set('marL', str(int(indent_in * E))); pPr.set('indent', str(int(-indent_in * E)))
        for tag in ('a:buClr', 'a:buChar', 'a:buNone'):
            for el in pPr.findall(qn(tag)): pPr.remove(el)
        bc = etree.SubElement(pPr, qn('a:buClr')); sc = etree.SubElement(bc, qn('a:srgbClr')); sc.set('val', color)
        ch = etree.SubElement(pPr, qn('a:buChar')); ch.set('char', char)


def figure(slide, path, x, y, w, h, align='center', name=None):
    """Fit an image inside the box without cropping."""
    im = Image.open(path); ar = im.width / im.height
    nw, nh = (h * ar, h) if w / h > ar else (w, w / ar)
    nx = x + (w - nw) / 2 if align == 'center' else x
    pic = slide.shapes.add_picture(path, Inches(nx), Inches(y + (h - nh) / 2), Inches(nw), Inches(nh))
    if name: pic.name = name
    return pic


def photo(slide, path, x, y, w, h, rounded=True, adj=5000, name=None):
    """Fill the box with a centre-cropped, EXIF-rotated, downsized copy of a photo (cached)."""
    im = ImageOps.exif_transpose(Image.open(path)).convert('RGB')
    ar_box, ar = w / h, im.width / im.height
    if ar > ar_box:
        nw = int(im.height * ar_box); x0 = (im.width - nw) // 2; im = im.crop((x0, 0, x0 + nw, im.height))
    else:
        nh = int(im.width / ar_box); y0 = (im.height - nh) // 2; im = im.crop((0, y0, im.width, y0 + nh))
    maxpx = int(max(w, h) * 220)
    if max(im.size) > maxpx:
        sc = maxpx / max(im.size); im = im.resize((int(im.width * sc), int(im.height * sc)), Image.LANCZOS)
    cache = THEME['cache'] or os.path.join(os.path.dirname(os.path.abspath(path)), '_pptx_cache')
    os.makedirs(cache, exist_ok=True)
    f = os.path.join(cache, f"{os.path.splitext(os.path.basename(path))[0]}_{w:.2f}x{h:.2f}.jpg"); im.save(f, quality=90)
    pic = slide.shapes.add_picture(f, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded:
        g = pic._element.spPr.find(qn('a:prstGeom')); g.set('prst', 'roundRect')
        av = g.find(qn('a:avLst'))
        if av is None: av = etree.SubElement(g, qn('a:avLst'))
        for c in list(av): av.remove(c)
        gd = etree.SubElement(av, qn('a:gd')); gd.set('name', 'adj'); gd.set('fmla', f'val {adj}')
    if name: pic.name = name
    return pic


def cover_photo(slide, path, x, y, w, h, focus=(0.5, 0.5), grad=None, dpi=150, name=None):
    """Edge-to-edge photo: crop to the box (cover) around `focus` (0-1, 0-1), no rounding.
    grad = (side, alpha_at_side, alpha_at_far_end, hex) lays a one-directional colour fade under the text side,
    e.g. ('left', 0.85, 0.0, '0A1322'); side is 'left' | 'right' | 'top' | 'bottom'."""
    im = ImageOps.exif_transpose(Image.open(path)).convert('RGB')
    tw, th = int(w * dpi), int(h * dpi); ar_box, ar = tw / th, im.width / im.height
    if ar > ar_box:
        nw = int(im.height * ar_box); x0 = int((im.width - nw) * focus[0]); im = im.crop((x0, 0, x0 + nw, im.height))
    else:
        nh = int(im.width / ar_box); y0 = int((im.height - nh) * focus[1]); im = im.crop((0, y0, im.width, y0 + nh))
    im = im.resize((tw, th), Image.LANCZOS)
    if grad:
        side, a0, a1, col = grad
        horiz = side in ('left', 'right'); n = tw if horiz else th
        mask = Image.new('L', im.size); d = ImageDraw.Draw(mask)
        for i in range(n):
            t = i / max(n - 1, 1)
            a = a0 + (a1 - a0) * t if side in ('left', 'top') else a1 + (a0 - a1) * t
            d.line([(i, 0), (i, th)] if horiz else [(0, i), (tw, i)], fill=int(255 * a))
        im = Image.composite(Image.new('RGB', im.size, tuple(int(col[k:k + 2], 16) for k in (0, 2, 4))), im, mask)
    cache = THEME['cache'] or os.path.join(os.path.dirname(os.path.abspath(path)), '_pptx_cache')
    os.makedirs(cache, exist_ok=True)
    f = os.path.join(cache, f"{os.path.splitext(os.path.basename(path))[0]}_cover_{w:.2f}x{h:.2f}_{focus[0]:.2f}_{focus[1]:.2f}_{grad[0] if grad else 'n'}.jpg")
    im.save(f, quality=90)
    pic = slide.shapes.add_picture(f, Inches(x), Inches(y), Inches(w), Inches(h))
    if name: pic.name = name
    return pic


def arrow(slide, x1, y1, x2, y2, color=None, w=1.5, dash=None, head=True):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color or THEME['grey']); c.line.width = Pt(w)
    if dash: c.line.dash_style = dash
    if head:
        ln = c.line._get_or_add_ln(); te = etree.SubElement(ln, qn('a:tailEnd'))
        te.set('type', 'triangle'); te.set('w', 'med'); te.set('len', 'med')
    return c


# ------------------------------------------------------------------ slide frame
def blank_slide(prs, bg=None):
    """Empty slide with an optional solid background colour (hex)."""
    s = prs.slides.add_slide(blank_layout(prs))
    if bg:
        s.background.fill.solid(); s.background.fill.fore_color.rgb = rgb(bg)
    return s


def content_slide(prs, tag, title, source=None, title_size=30):
    """Headline title (0.6, 0.6) / source line (0.6, 7.02) / slide number (12.2, 7.02).
    `tag` is an optional small section label above the title; pass '' to leave it out (recommended)."""
    s = prs.slides.add_slide(blank_layout(prs))
    if tag: text(s, 0.6, 0.32, 8, 0.3, tag.upper(), size=12, bold=True, color=THEME['accent'], name='Section tag')
    text(s, 0.6, 0.6, 12.1, 0.95, title, size=title_size, bold=True, color=THEME['navy'], name='Title')
    if source: text(s, 0.6, 7.02, 11.2, 0.3, source, size=10, color=THEME['grey'], name='Source')
    text(s, 12.2, 7.02, 0.55, 0.3, '0', size=10, color=THEME['grey'], align=PP_ALIGN.RIGHT, name='Slide number')
    return s


def renumber(prs):
    """Write 1..N into every 'Slide number' box (and digit boxes sitting at the slide-number spot)."""
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame: continue
            spot = abs(sh.left / E - 12.2) < 0.02 and abs(sh.top / E - 7.02) < 0.02 and sh.text_frame.text.strip().isdigit()
            if sh.name == 'Slide number' or spot: set_text_keep_style(sh, str(i))


# ------------------------------------------------------------------ editing an existing (user-edited) deck
def title_of(s):
    for sh in s.shapes:
        if sh.name == 'Title' and sh.has_text_frame: return sh.text_frame.text
    for sh in s.shapes:   # hand-made slides: first non-empty text box at the title height
        if sh.has_text_frame and abs(sh.top / E - 0.6) < 0.05 and sh.text_frame.text.strip(): return sh.text_frame.text
    return ''


def slide_by_title(prs, starts):
    for s in prs.slides:
        if title_of(s).startswith(starts): return s
    raise KeyError(f'no slide whose title starts with {starts!r}')


def find(s, name=None, starts=None):
    for sh in s.shapes:
        if name and sh.name == name: return sh
        if starts and sh.has_text_frame and sh.text_frame.text.strip().startswith(starts): return sh
    raise KeyError(f'{name or starts!r} not on slide "{title_of(s)[:40]}"')


def set_text_keep_style(sh, new):
    """Replace the text of a one-paragraph shape, keeping the first run's formatting."""
    p = sh.text_frame.paragraphs[0]; runs = p.runs
    runs[0].text = new
    for r in runs[1:]: r._r.getparent().remove(r._r)


def move(sh, x=None, y=None, w=None, h=None):
    if x is not None: sh.left = Inches(x)
    if y is not None: sh.top = Inches(y)
    if w is not None: sh.width = Inches(w)
    if h is not None: sh.height = Inches(h)


def remove(sh):
    sh._element.getparent().remove(sh._element)


# ------------------------------------------------------------------ speaker notes
def notes(prs, s, t):
    """Set the speaker notes. Some user-made notes masters give new slides no body placeholder;
    borrow one from an existing slide's notes page."""
    ns = s.notes_slide
    if ns.notes_text_frame is None:
        donor = next(x.notes_slide.notes_placeholder for x in prs.slides
                     if x.has_notes_slide and x.notes_slide.notes_placeholder is not None)
        ns.shapes._spTree.append(copy.deepcopy(donor._element))
    ns.notes_text_frame.text = t


def note_replace(s, old, new):
    """Replace one phrase in the notes; fails loudly if the phrase is not there (so edits never silently miss)."""
    tf = s.notes_slide.notes_text_frame
    assert old in tf.text, (title_of(s)[:30], old)
    tf.text = tf.text.replace(old, new, 1)


# ------------------------------------------------------------------ animation
def click_fade(slide, shapes, dur=500):
    """One on-click fade-in per shape, in order (raw p:timing). Group a row first if several shapes appear together.
    Mark each click in the notes with [클릭] so the speaker knows where to click."""
    P_ = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    assert slide._element.find(qn('p:timing')) is None, 'slide already has animations'
    n = [3]
    def nid():
        n[0] += 1; return n[0]
    clicks = ''
    for sh in shapes:
        a, b, c, d, e = nid(), nid(), nid(), nid(), nid()
        clicks += (f'<p:par><p:cTn id="{a}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                   f'<p:par><p:cTn id="{b}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                   f'<p:par><p:cTn id="{c}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">'
                   f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                   f'<p:set><p:cBhvr><p:cTn id="{d}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                   f'<p:tgtEl><p:spTgt spid="{sh.shape_id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                   f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                   f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{e}" dur="{dur}"/>'
                   f'<p:tgtEl><p:spTgt spid="{sh.shape_id}"/></p:tgtEl></p:cBhvr></p:animEffect>'
                   f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    xml = (f'<p:timing xmlns:p="{P_}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
           f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{clicks}</p:childTnLst></p:cTn>'
           f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
           f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')
    el = etree.fromstring(xml)
    ext = slide._element.find(qn('p:extLst'))
    if ext is not None: ext.addprevious(el)
    else: slide._element.append(el)


# ------------------------------------------------------------------ subsets
def keep_slides(prs, keep):
    """Drop every slide whose 1-based index is not in `keep` (used for 'send me only the changed slides')."""
    lst = prs.slides._sldIdLst
    for i, sid in reversed(list(enumerate(list(lst), 1))):
        if i not in keep:
            prs.part.drop_rel(sid.rId); lst.remove(sid)

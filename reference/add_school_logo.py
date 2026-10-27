# -*- coding: utf-8 -*-
"""
add_school_logo.py —— 只给已有 deck 贴校徽，不改封面、不换模板

用法：
    python add_school_logo.py --deck X.pptx --template TEMPLATE.pptx [--scale 0.55]
                              [--skip 1,-1] [--position top-right]

--skip 指定不贴校徽的页码（1 起，支持负索引），默认跳过第一页和最后一页。
模板中的校徽从版式图片读取，保留其 srcRect 裁剪。
"""
import argparse, copy, io
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

ap = argparse.ArgumentParser()
ap.add_argument("--deck", required=True)
ap.add_argument("--template", required=True)
ap.add_argument("--scale", type=float, default=0.55)
ap.add_argument("--margin", type=float, default=0.28, help="距右/上边缘（英寸）")
ap.add_argument("--skip", default="1,-1")
ap.add_argument("--only", default="", help="只给这些页码贴（1 起，逗号分隔）")
a = ap.parse_args()

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def find_logo(tpl):
    w = tpl.slide_width
    for master in tpl.slide_masters:
        for layout in master.slide_layouts:
            for sh in layout.shapes:
                if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.left > w * 0.55 and sh.top < 914400 * 1.5:
                    return sh
    return None


def remap(element, src_part, dst_slide):
    for el in element.iter():
        for attr, val in list(el.attrib.items()):
            if not attr.startswith(R_NS):
                continue
            rel = src_part.rels.get(val)
            if rel is None:
                continue
            try:
                if rel.reltype.endswith("/image"):
                    _, new_id = dst_slide.part.get_or_add_image_part(io.BytesIO(rel.target_part.blob))
                else:
                    new_id = dst_slide.part.relate_to(rel.target_ref, rel.reltype)
                el.set(attr, new_id)
            except Exception:
                pass


deck = Presentation(a.deck)
tpl = Presentation(a.template)
logo = find_logo(tpl)
if logo is None:
    raise SystemExit("模板里没找到校徽")

only = set()
for tok in a.only.split(","):
    tok = tok.strip()
    if tok:
        only.add(int(tok) - 1)

skip = set()
if only:
    skip = set(range(len(deck.slides))) - only
else:
    for tok in a.skip.split(","):
        tok = tok.strip()
        if not tok:
            continue
        i = int(tok)
        skip.add(i - 1 if i > 0 else len(deck.slides) + i)

SW = deck.slide_width
SH = deck.slide_height
margin = int(a.margin * 914400)

for i, slide in enumerate(deck.slides):
    if i in skip:
        continue
    el = copy.deepcopy(logo._element)
    remap(el, logo.part, slide)
    xfrm = el.find(".//{%s}xfrm" % NS_A)
    if xfrm is not None:
        off = xfrm.find("{%s}off" % NS_A)
        ext = xfrm.find("{%s}ext" % NS_A)
        if off is not None and ext is not None:
            cx, cy = int(ext.get("cx")), int(ext.get("cy"))
            ncx, ncy = int(cx * a.scale), int(cy * a.scale)
            ext.set("cx", str(ncx)); ext.set("cy", str(ncy))
            off.set("x", str(SW - margin - ncx))
            off.set("y", str(margin))
    slide.shapes._spTree.append(el)

deck.save(a.deck)
print("logo added to", len(deck.slides) - len(skip), "slides")

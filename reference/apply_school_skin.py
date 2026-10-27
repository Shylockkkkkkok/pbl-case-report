# -*- coding: utf-8 -*-
"""
apply_school_skin.py —— 给引擎生成的 PPT 套上西安交通大学模板外观

用法：
    python apply_school_skin.py --deck DECK.pptx --template TEMPLATE.pptx \
        --out OUT.pptx [--cover 1] [--thanks -1] [--no-logo]

做三件事：
  1. 用模板的封面页替换 deck 的第 1 页
  2. 用模板的致谢页替换 deck 的最后一页（自动识别含“谢谢”的页）
  3. 给每个内容页右上角贴模板的校徽，并加一条校色横线

不修改输入的 deck 与模板。复制的是形状 XML 与图片关系，跨文件安全。
"""
from __future__ import annotations
import argparse, copy, io, re, sys
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def iter_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes)
        else:
            yield sh


def remap_rels(element, src_part, dst_slide):
    """把克隆 XML 里的 r:embed / r:link / r:id 指向目标 slide 的关系。
    src_part 可以是 slide.part，也可以是 layout.part（校徽来自版式）。"""
    for el in element.iter():
        for attr, val in list(el.attrib.items()):
            if not attr.startswith(R_NS):
                continue
            rel = src_part.rels.get(val)
            if rel is None:
                continue
            try:
                if rel.reltype == RT.IMAGE:
                    _, new_id = dst_slide.part.get_or_add_image_part(io.BytesIO(rel.target_part.blob))
                else:
                    new_id = dst_slide.part.relate_to(rel.target_ref, rel.reltype)
                el.set(attr, new_id)
            except Exception:
                pass


def clone_slide(dst_prs, src_slide, layout):
    new = dst_prs.slides.add_slide(layout)
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)
    tree = new.shapes._spTree
    for shp in src_slide.shapes:
        el = copy.deepcopy(shp._element)
        remap_rels(el, src_slide.part, new)
        tree.append(el)
    return new


def move_slide(prs, from_index, to_index):
    lst = prs.slides._sldIdLst
    items = list(lst)
    item = items[from_index]
    lst.remove(item)
    lst.insert(to_index, item)


def delete_slide(prs, index):
    lst = prs.slides._sldIdLst
    item = list(lst)[index]
    prs.part.drop_rel(item.get(qn("r:id")))
    lst.remove(item)


def blank_layout(prs):
    for layout in prs.slide_masters[0].slide_layouts:
        if not any(sh.is_placeholder for sh in layout.shapes):
            return layout
    return prs.slide_masters[0].slide_layouts[-1]


def find_thanks_index(prs):
    for i in range(len(prs.slides) - 1, -1, -1):
        for sh in iter_shapes(prs.slides[i].shapes):
            if sh.has_text_frame and re.search(r"谢谢|Thanks|致谢", sh.text_frame.text):
                return i + 1  # 1-based
    return -1


def find_logo(template):
    """模板中右上角的小图片视为校徽。"""
    w = template.slide_width
    for master in template.slide_masters:
        for layout in master.slide_layouts:
            for sh in layout.shapes:
                if sh.shape_type != MSO_SHAPE_TYPE.PICTURE:
                    continue
                if sh.left > w * 0.55 and sh.top < Emu(int(1.5 * 914400)) and sh.width < w * 0.45:
                    return sh
    for slide in template.slides:
        for sh in slide.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.left > w * 0.55 and sh.top < Emu(int(1.5 * 914400)):
                return sh
    return None


def accent_color(template):
    """从模板主题取 accent1，取不到就用西交蓝。"""
    from pptx.dml.color import RGBColor
    try:
        master = template.slide_masters[0]
        xml = master.part.package.presentation_part.element.xml
        import re as _re
        m = _re.search(r"<a:accent1>.*?val=\"([0-9A-Fa-f]{6})\"", template.part.package.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme").blob.decode("utf-8"), _re.S)
        if m:
            return RGBColor.from_string(m.group(1).upper())
    except Exception:
        pass
    return RGBColor(0x00, 0xA5, 0xE0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--cover", type=int, default=1, help="模板封面页序号（1 起）")
    ap.add_argument("--thanks", type=int, default=-1, help="模板致谢页序号（1 起，-1=自动）")
    ap.add_argument("--no-logo", action="store_true")
    ap.add_argument("--logo-scale", type=float, default=0.6, help="校徽缩放倍数，默认 0.6")
    ap.add_argument("--cover-title", help="替换封面大标题（第二大的文字）")
    ap.add_argument("--cover-hide-university", action="store_true", help="删除封面最大的那行（通常是校名）")
    ap.add_argument("--presenter", help="替换“汇报人：”后的内容")
    ap.add_argument("--date", help="替换“日期：”后的内容")
    a = ap.parse_args()

    deck = Presentation(a.deck)
    tpl = Presentation(a.template)
    layout = blank_layout(deck)

    tpl_thanks = a.thanks if a.thanks > 0 else find_thanks_index(tpl)
    if tpl_thanks < 0:
        sys.exit("模板里找不到致谢页，请用 --thanks 指定")

    # 先克隆，后删除：避免删除后 add_slide 复用同名 part 导致包内重名
    orig_cover = 0
    orig_thanks = find_thanks_index(deck)          # 1 起；-1 表示没有
    orig_thanks = orig_thanks - 1 if orig_thanks > 0 else None

    cover_new = clone_slide(deck, tpl.slides[a.cover - 1], layout)
    thanks_new = clone_slide(deck, tpl.slides[tpl_thanks - 1], layout)

    victims = [i for i in (orig_cover, orig_thanks) if i is not None]
    for i in sorted(set(victims), reverse=True):
        delete_slide(deck, i)

    n = len(deck.slides)
    move_slide(deck, n - 2, 0)                     # 封面克隆 → 第 1 页
    # 致谢克隆原本就是最后一页，无需移动

    # 3) 校徽 + 校色线
    if not a.no_logo:
        logo = find_logo(tpl)
        if logo is None:
            print("⚠ 模板里没找到校徽，跳过")
        else:
            color = accent_color(tpl)
            for i, slide in enumerate(deck.slides):
                if i == 0:
                    continue
                # 克隆版式里的图片元素本身，保留 srcRect 裁剪与位置
                el = copy.deepcopy(logo._element)
                remap_rels(el, logo.part, slide)
                if a.logo_scale and abs(a.logo_scale - 1.0) > 1e-6:
                    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
                    xfrm = el.find(".//{%s}xfrm" % NS_A)
                    if xfrm is not None:
                        off = xfrm.find("{%s}off" % NS_A)
                        ext = xfrm.find("{%s}ext" % NS_A)
                        if off is not None and ext is not None:
                            x, y = int(off.get("x")), int(off.get("y"))
                            cx, cy = int(ext.get("cx")), int(ext.get("cy"))
                            ncx, ncy = int(cx * a.logo_scale), int(cy * a.logo_scale)
                            ext.set("cx", str(ncx)); ext.set("cy", str(ncy))
                            off.set("x", str(x + cx - ncx))
                            off.set("y", str(y + (cy - ncy) // 2))
                slide.shapes._spTree.append(el)
                # 页码若落在校徽下面，挪到右下角
                lx, ly = logo.left, logo.top
                rx, ry = lx + logo.width, ly + logo.height
                for shp in list(slide.shapes):
                    if not shp.has_text_frame or not shp.text_frame.text.strip().isdigit():
                        continue
                    if shp.left < rx and shp.left + shp.width > lx and shp.top < ry and shp.top + shp.height > ly:
                        shp.left = Emu(int(12.30 * 914400))
                        shp.top = Emu(int(6.95 * 914400))
                bar = slide.shapes.add_shape(1, Emu(int(0.62 * 914400)),
                                             Emu(int((deck.slide_height / 914400 - 0.055) * 914400)),
                                             Emu(int(12.1 * 914400)), Emu(int(0.04 * 914400)))
                bar.fill.solid(); bar.fill.fore_color.rgb = color
                bar.line.fill.background()
                bar.shadow.inherit = False

    # 4) 封面文字
    if a.cover_title or a.presenter or a.date:
        def texts(slide):
            out = []
            def walk(shapes):
                for sh in shapes:
                    if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                        walk(sh.shapes)
                    elif sh.has_text_frame and sh.text_frame.text.strip():
                        out.append(sh)
            walk(slide.shapes)
            return out

        def size_of(sh):
            vals = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
            return max(vals) if vals else 0

        def put(sh, text):
            ps = sh.text_frame.paragraphs
            if ps and ps[0].runs:
                ps[0].runs[0].text = text
                for r in ps[0].runs[1:]:
                    r.text = ""
            else:
                sh.text_frame.text = text

        cover_slide = deck.slides[0]
        items = sorted(texts(cover_slide), key=size_of, reverse=True)
        if a.cover_hide_university and items:
            items[0]._element.getparent().remove(items[0]._element)
            items = items[1:]
        if a.cover_title and items:
            put(items[0], a.cover_title)
        for sh in items:
            t = sh.text_frame.text.strip()
            if a.presenter and t.startswith("汇报人"):
                put(sh, "汇报人：" + a.presenter)
            elif a.date and ("日期" in t or t.replace(" ", "").startswith("日期")):
                put(sh, "日   期：" + a.date)

    deck.save(a.out)
    print("saved", a.out, len(deck.slides), "slides")


if __name__ == "__main__":
    main()

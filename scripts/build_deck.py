# -*- coding: utf-8 -*-
"""
build_deck.py — 商务汇报PPT模板驱动生成引擎

原理:商务模板的设计基本都放在"幻灯片"上(而非母版),所以采用
「捐赠页(donor slide)深拷贝 + 文本槽位替换」的方式装配新 PPT——
对任何自备 .pptx 模板通用,引擎不依赖特定模板库。

用法:
  # 1) 分析模板:列出每页类型与全部文本槽位(生成 outline 前必做)
  python build_deck.py analyze <模板文件> [--out slots.json]

  # 2) 按 outline 装配
  python build_deck.py build <模板文件> <outline.json> <输出.pptx>

  # 3) 校验输出(查残留占位文本)
  python build_deck.py verify <输出.pptx>

outline.json 格式:
{
  "pages": [
    {"donor": 1, "slots": {"1:0": "标题", "1:3": "副标题"}},
    {"donor": 2, "slots": {"2:0": "目录"}},
    {"donor": 5, "slots": {"5:g0/1": "要点一"}}
  ]
}
donor  = 模板中的原始页码(1 起);pages 数组顺序 = 输出 PPT 页序
slots  = {槽位id: 新文本},多行文本用 \n 分段;槽位 id 来自 analyze 输出
drop   = [槽位id...] 要删除的形状(模板自带 logo/水印/版权图形),id 同样来自 analyze
fit    = true 时,该页替换文本若明显变长,自动按比例缩字号(下限 0.6 倍)防溢出
"""
import json
import re
import sys
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- 基础工具

PLACEHOLDER_RE = re.compile(
    r"(点击|单击|请输入|请点击|添加|ADD|Add|XXXX|Lorem|lorem|Click|CLICK|Welcome|"
    r"WELCOME|XX月XX日|占位|待补|示例文字|文字说明|此处)", re.I)


def is_placeholder(text):
    return bool(PLACEHOLDER_RE.search(text or ""))


def walk_shapes(shapes, prefix=""):
    """按文档顺序遍历形状(含组内递归),yield (path, shape)。path 稳定,可作槽位 id。"""
    for idx, shp in enumerate(shapes):
        path = f"{prefix}{idx}"
        yield path, shp
        if shp.shape_type == 6:  # GROUP
            yield from walk_shapes(shp.shapes, prefix=path + "/")


def shape_text(shp):
    try:
        if shp.has_text_frame:
            return shp.text_frame.text
    except Exception:
        pass
    return ""


def max_font_size(shp):
    sizes = []
    try:
        for para in shp.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    sizes.append(run.font.size.pt)
    except Exception:
        pass
    return max(sizes) if sizes else None


def classify_slide(slide):
    texts = [t for _, s in walk_shapes(slide.shapes) for t in [shape_text(s)] if t.strip()]
    joined = " ".join(texts)
    if re.search(r"谢谢|感谢|THANKS|Thanks|thank", joined):
        return "end"
    if re.search(r"目\s*录|CONTENTS|Contents", joined):
        return "toc"
    if len(texts) <= 4:
        return "cover_or_section"
    return "content"


# ---------------------------------------------------------------- 幻灯片深拷贝

def duplicate_slide(prs, source):
    """把 source 幻灯片深拷贝为新页(含图片关系重接),返回新 slide。"""
    new_slide = prs.slides.add_slide(source.slide_layout)
    # 清掉 add_slide 从版式带进来的占位形状
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)
    # 深拷贝全部形状
    for shp in source.shapes:
        new_slide.shapes._spTree.append(deepcopy(shp._element))
    # 重接图片/媒体关系
    for rId, rel in source.part.rels.items():
        if rel.reltype in (RT.IMAGE, RT.MEDIA, RT.CHART) and not rel.is_external:
            new_rId = new_slide.part.relate_to(rel.target_part, rel.reltype)
            for el in new_slide.shapes._spTree.iter():
                for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                    if el.get(attr) == rId:
                        el.set(attr, new_rId)
    return new_slide


def delete_slides(prs, indexes):
    """删除指定 0 基索引的页(从后往前删)。"""
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    for i in sorted(indexes, reverse=True):
        el = ids[i]
        rId = el.get(qn("r:id"))
        prs.part.drop_rel(rId)
        sldIdLst.remove(el)


# ---------------------------------------------------------------- 文本替换

def _set_para_text(para, text):
    """把段落文本设为 text,保留首个 run 的格式,删掉其余 run。"""
    runs = para.runs
    if not runs:
        para.text = text  # 无 run 时直接写(用段落默认格式)
        return
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def scale_font(shp, ratio):
    """按比例缩小文本框架内所有带显式字号 run 的字号(下限 0.6 由调用方保证)。"""
    try:
        for para in shp.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    run.font.size = int(run.font.size.pt * ratio + 0.5)
    except Exception:
        pass


def replace_text(shp, value):
    """把形状文本替换为 value;\n 分段,逐段沿用首段格式。"""
    tf = shp.text_frame
    paras = tf.paragraphs
    if not paras:
        return
    lines = str(value).split("\n")
    _set_para_text(paras[0], lines[0])
    # 删除多余段落
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    # 需要更多段落时,克隆首段
    if len(lines) > 1:
        first = tf.paragraphs[0]._p
        for line in lines[1:]:
            clone = deepcopy(first)
            first.addnext(clone)
            first = clone
            from pptx.text.text import _Paragraph
            _set_para_text(_Paragraph(clone, tf), line)


# ---------------------------------------------------------------- analyze

def cmd_analyze(template, out=None, with_pics=False):
    prs = Presentation(str(template))
    result = {"template": str(template), "slides": len(prs.slides), "pages": []}
    for i, slide in enumerate(prs.slides, 1):
        page = {"slide": i, "type": classify_slide(slide), "slots": [], "pics": []}
        for path, shp in walk_shapes(slide.shapes):
            txt = shape_text(shp)
            if shp.shape_type == 13:  # PICTURE
                page["pics"].append(f"{i}:{path}")
                continue
            if not txt.strip():
                continue
            page["slots"].append({
                "id": f"{i}:{path}",
                "text": txt.replace("\n", "\\n")[:80],
                "placeholder": is_placeholder(txt),
                "size": max_font_size(shp),
            })
        result["pages"].append(page)
    js = json.dumps(result, ensure_ascii=False, indent=1)
    if out:
        Path(out).write_text(js, encoding="utf-8")
        print(f"analyze -> {out} ({result['slides']} slides)")
    else:
        print(js)
    return result


# ---------------------------------------------------------------- build

def cmd_build(template, outline_path, output):
    outline = json.loads(Path(outline_path).read_text(encoding="utf-8"))
    prs = Presentation(str(template))
    n_orig = len(prs.slides)
    pages = outline["pages"]

    for entry in pages:
        donor_no = entry["donor"]
        if not (1 <= donor_no <= n_orig):
            raise SystemExit(f"outline donor {donor_no} 越界(模板共 {n_orig} 页)")
        source = prs.slides[donor_no - 1]
        new_slide = duplicate_slide(prs, source)
        # 槽位替换:按 (页码:路径) 找形状
        wanted = entry.get("slots", {})
        fit = bool(entry.get("fit"))
        applied = set()
        for path, shp in walk_shapes(new_slide.shapes):
            sid = f"{donor_no}:{path}"
            if sid in wanted:
                old_text = shape_text(shp)
                value = str(wanted[sid])
                replace_text(shp, value)
                applied.add(sid)
                # fit: 新文本明显变长时按比例缩字号,防溢出出界(下限 0.6 倍)
                if fit:
                    old_len = sum(2 if ord(c) > 127 else 1 for c in old_text)
                    new_len = sum(2 if ord(c) > 127 else 1 for c in value)
                    if old_len > 0 and new_len > old_len * 1.05:
                        scale_font(shp, max(0.6, old_len / new_len))
        missing = set(wanted) - applied
        if missing:
            print(f"[warn] donor {donor_no} 未命中槽位: {sorted(missing)}")
        # 形状删除(模板自带 logo/水印/版权图形)
        drops = set(entry.get("drop", []))
        if drops:
            victims = [(p, s) for p, s in walk_shapes(new_slide.shapes)
                       if f"{donor_no}:{p}" in drops]
            for p, s in victims:
                s._element.getparent().remove(s._element)
            gone = {f"{donor_no}:{p}" for p, _ in victims}
            if drops - gone:
                print(f"[warn] donor {donor_no} drop 未命中: {sorted(drops - gone)}")
    delete_slides(prs, range(n_orig))
    prs.save(str(output))
    print(f"build -> {output} ({len(pages)} pages)")


# ---------------------------------------------------------------- verify

def cmd_verify(pptx_path):
    prs = Presentation(str(pptx_path))
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        for path, shp in walk_shapes(slide.shapes):
            t = shape_text(shp)
            if t.strip() and is_placeholder(t):
                bad.append((i, f"{i}:{path}", t.strip()[:40]))
    if bad:
        print(f"[verify] 残留占位文本 {len(bad)} 处:")
        for pg, sid, t in bad:
            print(f"   p{pg:02d} {sid}: {t}")
        return 1
    print(f"[verify] OK,无残留占位文本({len(prs.slides)} pages)")
    return 0


# ---------------------------------------------------------------- CLI

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    cmd = sys.argv[1]
    if cmd == "analyze":
        out = None
        if "--out" in sys.argv:
            out = sys.argv[sys.argv.index("--out") + 1]
        cmd_analyze(sys.argv[2], out)
    elif cmd == "build":
        cmd_build(sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == "verify":
        sys.exit(cmd_verify(sys.argv[2]))
    else:
        print(__doc__)
        sys.exit(1)

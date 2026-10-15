# -*- coding: utf-8 -*-
"""
make_demo_template.py — 生成一个原创的合成演示模板(assets/demo_template.pptx)

用途:
  1. 发布版技能的自包含示例:没有自备模板库时,用本模板跑通 analyze→build→verify→preview 全链路;
  2. 示范一个"好模板"的结构:封面/目录/过渡/卡片内容/数据/结尾 六类页型,文本全部放在
     独立文本框里(analyze 才能认出槽位),设计色直接写在形状上。

设计:16:9、白底、微软雅黑、商务红黑强调色(#BE1007)——规律来自 91 套欧美扁平风模板实测
(见 references/设计系统.md),但图形全部由本脚本从零绘制,不含任何第三方模板素材。

用法:
  python make_demo_template.py [输出路径]     # 默认 ../assets/demo_template.pptx
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- 设计常量
# 三套主题调色板:规律来自 91 套欧美扁平风模板实测(见 references/设计系统.md),
# 图形全部由本脚本从零绘制,不含任何第三方模板素材。
THEMES = {
    # 商务红黑:正式年终汇报/党政机关
    "redblack": dict(ACCENT="BE1007", DARK="202022", GRAY="868687",
                     FAINT="9A9A9A", LIGHT="F7F7F6", LINE="E3E3E3", DECO="F1F1F0"),
    # 蓝色商务:项目方案/投标答辩/科技向
    "blue":     dict(ACCENT="1F5FA8", DARK="1F2A3D", GRAY="6B7686",
                     FAINT="98A2AE", LIGHT="EEF3F9", LINE="D5DEE8", DECO="E4EDF7"),
    # 清新文艺:产品宣传/企业文化/年轻向
    "fresh":    dict(ACCENT="0F9E8E", DARK="26302B", GRAY="6E7B74",
                     FAINT="9AA5A0", LIGHT="F0F7F4", LINE="D8E5DF", DECO="E6F2EE"),
}

ACCENT = "BE1007"   # 强调红(商务红黑族,默认主题)
DARK = "202022"     # 主文字
GRAY = "868687"     # 次要文字
FAINT = "9A9A9A"    # 页脚/口径
LIGHT = "F7F7F6"    # 卡片/装饰面
LINE = "E3E3E3"     # 卡片描边
DECO = "F1F1F0"     # 大装饰面
FONT = "微软雅黑"


def set_theme(name):
    """切换主题(重绑定模块级颜色常量;页型函数在生成时读取)。"""
    global ACCENT, DARK, GRAY, FAINT, LIGHT, LINE, DECO
    t = THEMES[name]
    ACCENT, DARK, GRAY, FAINT = t["ACCENT"], t["DARK"], t["GRAY"], t["FAINT"]
    LIGHT, LINE, DECO = t["LIGHT"], t["LINE"], t["DECO"]

PAGE_W, PAGE_H = Inches(13.333), Inches(7.5)


def _style_run(run, size, color, bold=False):
    """设置 run 字体:拉丁 + 东亚都指向微软雅黑(只设 font.name 不管汉字)。"""
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", FONT)


def add_text(slide, x, y, w, h, text, size, color, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=None):
    """加一个文本框;返回形状。多行文本用 \\n 分段。"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing:
            p.line_spacing = spacing
        run = p.add_run()
        run.text = line
        _style_run(run, size, color, bold)
    return box


def add_rect(slide, x, y, w, h, fill, line=None, rounded=False):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded:
        shp.adjustments[0] = 0.06
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line:
        shp.line.color.rgb = RGBColor.from_string(line)
        shp.line.width = Pt(0.75)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_oval(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor.from_string(fill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def page_title(slide, text):
    """内容页页标题 + 强调色短杠(全库统一构图)。"""
    add_text(slide, 0.9, 0.72, 9.5, 0.6, text, 28, DARK, bold=True)
    add_rect(slide, 0.92, 1.44, 0.9, 0.05, ACCENT)


def footer(slide, text):
    add_text(slide, 0.9, 6.92, 8.0, 0.3, text, 9, FAINT)


# ---------------------------------------------------------------- 六类页型

def slide_cover(prs):
    """封面:装饰圆 + 左侧强调竖杠 + 主标/副标/汇报信息 + 底部色带。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_oval(s, 10.6, -2.0, 4.6, 4.6, DECO)
    add_rect(s, 0.9, 2.62, 0.12, 1.95, ACCENT)
    add_text(s, 1.25, 2.55, 8.2, 1.15, "季度业务汇报", 54, DARK, bold=True)
    add_text(s, 1.25, 3.82, 8.2, 0.5, "2026年第三季度经营分析", 18, GRAY)
    add_text(s, 1.25, 4.55, 8.2, 0.85, "汇报人：张明\n日期：2026年10月", 12, GRAY, spacing=1.3)
    add_rect(s, 0, 7.22, 13.333, 0.28, ACCENT)
    return s


def slide_toc(prs):
    """目录:编号(强调色)+ 条目(深色),右侧浅色装饰面板。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_text(s, 0.9, 0.78, 4.0, 0.65, "目录", 32, DARK, bold=True)
    add_rect(s, 0.92, 1.56, 0.9, 0.05, ACCENT)
    add_rect(s, 9.8, 0, 3.533, 7.5, LIGHT)
    add_rect(s, 10.35, 3.55, 0.55, 0.07, ACCENT)
    entries = [("01", "经营业绩回顾"), ("02", "关键指标分析"), ("03", "下季度工作计划")]
    for i, (num, name) in enumerate(entries):
        y = 2.25 + i * 1.32
        add_text(s, 0.95, y, 0.9, 0.55, num, 24, ACCENT, bold=True)
        add_text(s, 1.95, y + 0.03, 6.5, 0.55, name, 20, DARK)
    return s


def slide_section(prs):
    """过渡页:大编号 + 章节名 + 英文小字 + 底部色带。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_text(s, 0.9, 2.15, 3.2, 1.6, "01", 80, ACCENT, bold=True)
    add_text(s, 0.95, 3.88, 9.0, 0.9, "经营业绩回顾", 36, DARK, bold=True)
    add_text(s, 0.95, 4.85, 9.0, 0.4, "BUSINESS REVIEW", 13, GRAY)
    add_rect(s, 0, 7.22, 13.333, 0.28, ACCENT)
    return s


def _card(slide, x, title, body):
    """一张卡片:浅色圆角矩形 + 强调小杠 + 卡片标题 + 正文。"""
    add_rect(slide, x, 2.2, 3.6, 3.7, LIGHT, line=LINE, rounded=True)
    add_rect(slide, x + 0.35, 2.62, 0.5, 0.07, ACCENT)
    add_text(slide, x + 0.35, 2.88, 2.95, 0.5, title, 18, DARK, bold=True)
    add_text(slide, x + 0.35, 3.52, 2.95, 2.0, body, 12, GRAY, spacing=1.3)


def slide_cards(prs):
    """内容页(卡片构图):页标题 + 三张并列卡片 + 页脚。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    page_title(s, "本季度经营概况")
    _card(s, 0.9, "营收规模", "营业收入18.6亿元\n同比增长18%\n创单季历史新高")
    _card(s, 4.85, "盈利水平", "毛利率34.2%\n环比提升1.1个百分点\n三项费用率下降")
    _card(s, 8.8, "市场覆盖", "新增签约客户46家\n覆盖华东华南主要区域\n渠道伙伴达120家")
    footer(s, "2026Q3 · 经营分析")
    return s


def slide_data(prs):
    """数据页:左侧大数字 + 右侧三行指标 + 口径脚注。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    page_title(s, "关键指标一览")
    add_text(s, 0.95, 2.35, 4.0, 1.3, "18.6%", 60, ACCENT, bold=True)
    add_text(s, 0.95, 3.75, 4.0, 0.4, "营收同比增长", 14, GRAY)
    add_rect(s, 5.15, 2.4, 0.012, 2.9, LINE)
    metrics = [("34.2%", "毛利率,环比+1.1pct"),
               ("46家", "新增签约客户"),
               ("91%", "重点客户续约率")]
    for i, (num, label) in enumerate(metrics):
        y = 2.42 + i * 1.05
        add_text(s, 5.55, y, 1.9, 0.5, num, 24, DARK, bold=True)
        add_text(s, 7.55, y + 0.08, 4.6, 0.4, label, 12, GRAY)
    footer(s, "口径:2026Q3 vs 2025Q3,未经审计")
    return s


def slide_end(prs):
    """结尾页:致谢 + 英文注 + 底部色带。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 6.27, 2.52, 0.8, 0.07, ACCENT)
    add_text(s, 0.9, 2.85, 11.53, 1.0, "谢谢观看", 44, DARK, bold=True, align=PP_ALIGN.CENTER)
    add_text(s, 0.9, 3.95, 11.53, 0.4, "THANKS FOR WATCHING", 13, GRAY, align=PP_ALIGN.CENTER)
    add_rect(s, 0, 7.22, 13.333, 0.28, ACCENT)
    return s


def main():
    args = sys.argv[1:]
    style = "redblack"
    if "--style" in args:
        i = args.index("--style")
        style = args[i + 1]
        del args[i:i + 2]
    if style not in THEMES:
        raise SystemExit(f"未知主题 {style!r},可选: {list(THEMES)}")
    set_theme(style)

    out = Path(args[0]) if args else \
        Path(__file__).resolve().parent.parent / "assets" / "demo_template.pptx"
    out.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width, prs.slide_height = PAGE_W, PAGE_H
    slide_cover(prs)
    slide_toc(prs)
    slide_section(prs)
    slide_cards(prs)
    slide_data(prs)
    slide_end(prs)
    prs.save(str(out))
    print(f"demo template ({style}) -> {out} ({len(prs.slides._sldIdLst)} slides)")


if __name__ == "__main__":
    main()

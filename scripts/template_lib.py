# -*- coding: utf-8 -*-
"""
template_lib.py — 商务模板库检索/校验工具(BYO:模板文件用户自备)

用法:
  python template_lib.py families            # 风格族列表 + 每族数量
  python template_lib.py list [--family X] [--ooxml-only] [--dark]
  python template_lib.py pick <风格族>       # 推荐首选(优先 OOXML 可直接构建的)
  python template_lib.py path <编号或文件名> # 返回模板路径
  python template_lib.py check               # 模板库配置状态

模板库根目录解析顺序:
  1. 环境变量 BUSINESS_PPT_ROOT(推荐,换机器不用改文件)
  2. scripts/templates.json 的 root 字段(非空且目录存在时)
  3. 都未配置 → BYO 模式:索引数据仍可浏览(families/list),pick/path 只给建议;
     构建不受影响——build_deck.py 对任何 .pptx 都能跑,新手可先用
     技能自带的 assets/demo_template.pptx 跑通全链路。
"""
import json
import os
import sys
from pathlib import Path

DATA = Path(__file__).resolve().parent / "templates.json"

BYO_HINT = (
    "模板库未配置(BYO 模式):\n"
    "  自备模板目录后,设环境变量 BUSINESS_PPT_ROOT 指向它,\n"
    "  或填 scripts/templates.json 的 root 字段。\n"
    "  未配置也能构建:对任何 .pptx 跑 build_deck.py analyze/build,\n"
    "  或先用技能自带的 assets/demo_template.pptx 跑通全链路。"
)


def resolve_root(d):
    """返回 (root, 来源说明);未配置返回 (None, None)。"""
    env = os.environ.get("BUSINESS_PPT_ROOT")
    if env:
        return env, "环境变量 BUSINESS_PPT_ROOT"
    root = (d.get("root") or "").strip()
    if root and Path(root).is_dir():
        return root, "templates.json 的 root 字段"
    return None, None


def load():
    d = json.loads(DATA.read_text(encoding="utf-8"))
    root, src = resolve_root(d)
    d["_root"], d["_root_src"] = root, src
    for t in d["templates"]:
        t["path"] = str(Path(root) / t["file"]) if root else ""
        t["exists"] = bool(root) and Path(t["path"]).exists()
    return d


def cmd_families(d):
    from collections import Counter
    c = Counter(t["family"] for t in d["templates"])
    print(f"{'风格族':10s} {'数量':>4s}  OOXML可用")
    for fam, n in c.most_common():
        ok = sum(1 for t in d["templates"] if t["family"] == fam and t["fmt"] == "pptx")
        print(f"{fam:10s} {n:4d}  {ok}")


def cmd_list(d, family=None, ooxml_only=False, dark=None):
    rows = [t for t in d["templates"]
            if (not family or t["family"] == family)
            and (not ooxml_only or t["fmt"] == "pptx")
            and (dark is None or t["dark"] == dark)]
    print(f"{'编号':>4s} {'文件':34s} {'风格族':10s} {'格式':6s} 主色")
    for t in rows:
        pal = "/".join(t["palette"][:3])
        print(f"{t['num']:>4d} {t['file']:34s} {t['family']:10s} {t['fmt']:6s} {pal}")
    print(f"-- {len(rows)} 套")


def cmd_pick(d, family):
    cands = [t for t in d["templates"] if t["family"] == family]
    if not cands:
        print(f"没有风格族 {family!r}"); sys.exit(1)
    cands.sort(key=lambda t: (t["fmt"] != "pptx", t["num"]))  # OOXML 优先,再按编号
    for t in cands[:3]:
        mark = "★首选" if t is cands[0] else " 备选"
        print(f"{mark} {t['num']:03d} {t['file']}  [{t['fmt']}]")
    if d["_root"]:
        print(f"path: {cands[0]['path']}")
    else:
        print("path: (模板库未配置——设 BUSINESS_PPT_ROOT 后此命令返回实际路径)")


def cmd_path(d, key):
    for t in d["templates"]:
        if key in (str(t["num"]), t["file"], t["file"].rsplit(".", 1)[0]):
            if d["_root"]:
                print(t["path"] if t["exists"] else f"[缺失] {t['path']}")
            else:
                print(f"[模板库未配置] 目标文件: {t['file']};设 BUSINESS_PPT_ROOT 后可得实际路径")
            return
    print(f"未找到 {key!r}"); sys.exit(1)


def cmd_check(d):
    root, src = d["_root"], d["_root_src"]
    if not root:
        print(BYO_HINT)
        print(f"内置索引仍可浏览:共 {len(d['templates'])} 套第三方模板包的风格族/主色板元数据。")
        return
    miss = [t for t in d["templates"] if not t["exists"]]
    ok = [t for t in d["templates"] if t["exists"]]
    ooxml = [t for t in ok if t["fmt"] == "pptx"]
    print(f"模板库 root: {root}  (来源: {src})")
    print(f"共 {len(d['templates'])} 套 | 存在 {len(ok)} | 缺失 {len(miss)}")
    print(f"可直接构建(OOXML): {len(ooxml)} 套;旧二进制 .ppt 需先另存为 .pptx: {len(ok) - len(ooxml)} 套")
    if miss:
        print("缺失:", ", ".join(t["file"] for t in miss[:10]))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    d = load()
    cmd, args = sys.argv[1], sys.argv[2:]
    if cmd == "families":
        cmd_families(d)
    elif cmd == "list":
        fam = args[args.index("--family") + 1] if "--family" in args else None
        cmd_list(d, fam, "--ooxml-only" in args,
                 True if "--dark" in args else (False if "--light" in args else None))
    elif cmd == "pick":
        cmd_pick(d, args[0])
    elif cmd == "path":
        cmd_path(d, args[0])
    elif cmd == "check":
        cmd_check(d)
    else:
        print(__doc__); sys.exit(1)

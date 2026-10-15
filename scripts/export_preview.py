# -*- coding: utf-8 -*-
"""
export_preview.py — 用本机 PowerPoint 把 PPTX 每页导出为 PNG,供逐页目检。

用法:
  python export_preview.py <pptx> <输出目录> [宽] [高]
  默认 1920x1080。

依赖: pywin32 + 本机安装 Microsoft PowerPoint。
注意: 导出期间不要手动操作 PowerPoint;脚本结束会自动关闭演示文稿。
"""
import os
import sys


def export(pptx_path, outdir, width=1920, height=1080):
    import win32com.client  # 延迟导入,便于无 PowerPoint 环境给出清晰报错

    pptx_path = os.path.abspath(pptx_path)
    outdir = os.path.abspath(outdir)
    os.makedirs(outdir, exist_ok=True)

    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(pptx_path, WithWindow=False)
        try:
            n = pres.Slides.Count
            for i in range(1, n + 1):
                out = os.path.join(outdir, f"slide_{i:02d}.png")
                pres.Slides(i).Export(out, "PNG", width, height)
                print(f"  p{i:02d} -> {out}")
            print(f"exported {n} slides -> {outdir}")
        finally:
            pres.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    w = int(sys.argv[3]) if len(sys.argv) > 3 else 1920
    h = int(sys.argv[4]) if len(sys.argv) > 4 else 1080
    export(sys.argv[1], sys.argv[2], w, h)

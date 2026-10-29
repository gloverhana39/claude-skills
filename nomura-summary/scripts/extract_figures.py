# -*- coding: utf-8 -*-
"""原資料PDFから写真と図表を切り出す。

python3 extract_figures.py <資料.pdf> <出力dir> [--dpi 250]

出力:
  photo_pXX_N.jpg   … PDFに埋め込まれた写真（幅300px以上のJPEG/RGB）
  fig_pXX_N.png     … 「図表」「（出所）」の位置から推定した図表の切り出し
  figures.tsv       … 切り出し一覧（ファイル名・ページ・座標・見出し）
  contact.jpg       … 一覧の縮小画像（目視確認用。必ず見て、ずれていれば座標を直して切り直す）

座標はPDFのポイント単位。手動で切り直すときは crop(pdf, page, x0, y0, x1, y1, out) を使う。
"""
import os
import re
import subprocess
import sys
import glob
from PIL import Image

CAPTION = re.compile(r"^(図表\s*\d+(-\d+)?|図\s*\d+)")


def run(cmd):
    return subprocess.run(cmd, check=True, capture_output=True, text=True).stdout


def render(pdf, page, dpi, tmp):
    base = os.path.join(tmp, f"r{page}")
    if not glob.glob(base + "-*.png"):
        run(["pdftoppm", "-f", str(page), "-l", str(page), "-r", str(dpi), "-png", pdf, base])
    return Image.open(glob.glob(base + "-*.png")[0])


def crop(pdf, page, x0, y0, x1, y1, out, dpi=250, tmp=None):
    tmp = tmp or os.path.dirname(out)
    im = render(pdf, page, dpi, tmp)
    s = dpi / 72
    im.crop((int(x0 * s), int(y0 * s), int(x1 * s), int(y1 * s))).save(out)


def lines_by_page(pdf, tmp):
    html = os.path.join(tmp, "bbox.html")
    run(["pdftotext", "-bbox-layout", pdf, html])
    t = open(html, encoding="utf8").read()
    pages = []
    for pg in t.split("<page ")[1:]:
        w = float(re.search(r'width="([\d.]+)"', pg).group(1))
        h = float(re.search(r'height="([\d.]+)"', pg).group(1))
        ls = []
        for x0, y0, x1, y1, body in re.findall(
                r'<line xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</line>', pg, re.S):
            txt = "".join(re.findall(r">([^<]*)</word>", body))
            ls.append((float(x0), float(y0), float(x1), float(y1), txt))
        pages.append((w, h, ls))
    return pages


def main():
    pdf, outdir = sys.argv[1], sys.argv[2]
    dpi = int(sys.argv[sys.argv.index("--dpi") + 1]) if "--dpi" in sys.argv else 250
    os.makedirs(outdir, exist_ok=True)
    tmp = os.path.join(outdir, "_tmp")
    os.makedirs(tmp, exist_ok=True)
    rows = []
    # 写真
    run(["pdfimages", "-j", pdf, os.path.join(tmp, "img")])
    lst = run(["pdfimages", "-list", pdf]).splitlines()[2:]
    for ln in lst:
        c = ln.split()
        page, num, typ, w, h, color, enc = int(c[0]), int(c[1]), c[2], int(c[3]), int(c[4]), c[5], c[8]
        if typ == "image" and enc == "jpeg" and color in ("rgb", "icc") and w >= 300 and w / h < 3:  # 横長の装飾帯は除外
            src = os.path.join(tmp, f"img-{num:03d}.jpg")
            if os.path.exists(src):
                dst = f"photo_p{page:02d}_{num}.jpg"
                os.replace(src, os.path.join(outdir, dst))
                rows.append((dst, page, "", f"写真 {w}x{h}"))
    # 図表
    for pi, (pw, ph, ls) in enumerate(lines_by_page(pdf, tmp), 1):
        caps = [l for l in ls if CAPTION.match(l[4])]
        subs = [c for c in caps if re.match(r"^図表\s*\d+-\d+", c[4])]
        use = subs if subs else caps
        srcs = [l for l in ls if l[4].startswith("（出所）")]
        for k, (x0, y0, x1, y1, txt) in enumerate(sorted(use, key=lambda c: (c[1] // 20, c[0]))):
            left = x0 < pw / 2
            if subs:  # 左右2枚組
                cx0, cx1 = (18, pw / 2) if left else (pw / 2, pw - 15)
            else:
                cx0, cx1 = 18, pw - 15
            same_col = [s for s in srcs if s[1] > y1 and ((s[0] < pw / 2) == left or not subs)]
            yb = max([s[3] for s in same_col], default=min(ph - 30, y1 + 300))
            name = f"fig_p{pi:02d}_{k + 1}.png"
            crop(pdf, pi, cx0, y0 - 4, cx1, yb + 4, os.path.join(outdir, name), dpi, tmp)
            rows.append((name, pi, f"{cx0:.0f},{y0 - 4:.0f},{cx1:.0f},{yb + 4:.0f}", txt[:40]))
    with open(os.path.join(outdir, "figures.tsv"), "w", encoding="utf8") as f:
        f.write("file\tpdf_page\tcrop_pt(x0,y0,x1,y1)\tlabel\n")
        for r in rows:
            f.write("\t".join(map(str, r)) + "\n")
    # 目視確認用の一覧
    th = []
    for r in rows:
        im = Image.open(os.path.join(outdir, r[0])).convert("RGB")
        im.thumbnail((300, 240))
        th.append(im)
    if th:
        n = 5
        g = Image.new("RGB", (n * 310, ((len(th) + n - 1) // n) * 250), "white")
        for i, im in enumerate(th):
            g.paste(im, ((i % n) * 310, (i // n) * 250))
        g.save(os.path.join(outdir, "contact.jpg"))
    print(f"{len(rows)} items -> {outdir}/figures.tsv")


if __name__ == "__main__":
    main()

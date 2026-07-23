#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AsIs スクリーンショット注釈スクリプト
- materials/screenshots/ の実スクショから該当箇所をcrop
- 赤枠(問題)/青枠(良い点=Good)+番号バッジを描画
- 出力: deliverables/1784786903168-17720/assets/asis_*.png
ラベル文言そのものは pptx側でネイティブテキスト(凡例)として付与する
(スクショに焼き込むのはバッジ番号/Goodのみ。理由: 判読性と、後工程の文言調整を容易にするため)
"""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = "/workspace/akkodis-marketing-portal"
SRC = os.path.join(ROOT, "materials/screenshots")
OUT = os.path.join(ROOT, "deliverables/1784786903168-17720/assets")
os.makedirs(OUT, exist_ok=True)

RED = (196, 30, 30)
BLUE = (0, 90, 181)
WHITE = (255, 255, 255)

FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"


def font(size):
    return ImageFont.truetype(FONT_PATH, size)


def draw_box(draw, xyxy, color, halo_w=8, line_w=5):
    """白ハロー付きの矩形枠(コントラスト確保・色弱配慮)"""
    x0, y0, x1, y1 = xyxy
    draw.rectangle([x0 - halo_w // 2, y0 - halo_w // 2, x1 + halo_w // 2, y1 + halo_w // 2],
                   outline=WHITE, width=halo_w)
    draw.rectangle([x0, y0, x1, y1], outline=color, width=line_w)


def draw_number_badge(im, draw, center, number, color, size=44):
    cx, cy = center
    r = size // 2
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=color, outline=WHITE, width=4)
    f = font(int(size * 0.55))
    text = str(number)
    bbox = draw.textbbox((0, 0), text, font=f)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text((cx - tw / 2 - bbox[0], cy - th / 2 - bbox[1]), text, fill=WHITE, font=f)


def draw_good_badge(im, draw, topleft, size=30):
    x, y = topleft
    f = font(size)
    text = "Good"
    bbox = draw.textbbox((0, 0), text, font=f)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pad = 10
    w, h = tw + pad * 2, th + pad * 2
    draw.rounded_rectangle([x, y, x + w, y + h], radius=8, fill=BLUE, outline=WHITE, width=3)
    draw.text((x + pad - bbox[0], y + pad - bbox[1]), text, fill=WHITE, font=f)
    return (x, y, x + w, y + h)


def open_src(name):
    return Image.open(os.path.join(SRC, name)).convert("RGB")


def save(im, name):
    path = os.path.join(OUT, name)
    im.save(path, optimize=True)
    print("saved", path, im.size)


def annotate(im, boxes):
    """boxes: list of dict(xyxy, color, number|None, good=False, badge='tl'|'bl'|'tr')"""
    draw = ImageDraw.Draw(im)
    for b in boxes:
        color = BLUE if b.get("good") else RED
        draw_box(draw, b["xyxy"], color)
        x0, y0, x1, y1 = b["xyxy"]
        pos = b.get("badge", "tl")
        if pos == "bl":
            bx, by = x0, y1
        elif pos == "tr":
            bx, by = x1, y0
        else:
            bx, by = x0, y0
        if b.get("good"):
            draw_good_badge(im, draw, (x0 - 4, max(0, y0 - 46)))
            if b.get("number"):
                draw_number_badge(im, draw, (x0 + 6, max(24, y0 - 4)), b["number"], BLUE, size=34)
        else:
            draw_number_badge(im, draw, (bx, by), b["number"], RED, size=44)
    return im


# ---------------------------------------------------------------------------
# 1. トップ (top_pc / top_sp)
# ---------------------------------------------------------------------------
def build_top():
    im = open_src("top_pc.png")
    boxes = [
        dict(xyxy=(160, 410, 1375, 626), color=RED, number=1),   # T-3 4カード
        dict(xyxy=(568, 898, 954, 953), color=RED, number=2),    # T-4 業務依頼フォームボタン
        dict(xyxy=(160, 1012, 780, 1046), color=RED, number=2),  # T-4 ブランドレビューの手順リンク
        dict(xyxy=(80, 2630, 1385, 3200), color=RED, number=3),  # T-2 組織体制
    ]
    annotate(im, boxes)
    save(im, "asis_top_pc.png")

    sp = open_src("top_sp.png")
    crop = sp.crop((0, 2080, 458, 2720))
    boxes_sp = [dict(xyxy=(15, 2250 - 2080, 440, 2695 - 2080), color=RED, number=4)]
    annotate(crop, boxes_sp)
    save(crop, "asis_top_sp_dashboard.png")


# ---------------------------------------------------------------------------
# 2. ブランドライブラリ (brand_pc / brand_sp) — 最重要
# ---------------------------------------------------------------------------
def build_brand():
    im = open_src("brand_pc.png")

    # Crop A: 0-1400 (黄色マーカー3見出し ①B-6 / アイコンカード3枚 ②B-5)
    cropA = im.crop((0, 0, 1463, 1400))
    boxesA = [
        dict(xyxy=(160, 465, 1300, 945), color=RED, number=1),
        dict(xyxy=(155, 1108, 1310, 1310), color=RED, number=2),
    ]
    annotate(cropA, boxesA)
    save(cropA, "asis_brand_A.png")

    # Crop B: 1300-1820 (ガイドライン10本 ③B-2 / 表記ゆれ2行 ④B-4)
    cropB = im.crop((0, 1300, 1463, 1820))
    boxesB = [
        dict(xyxy=(185, 1558 - 1300, 1360, 1795 - 1300), color=RED, number=3, badge="tr"),
        dict(xyxy=(190, 1558 - 1300, 700, 1598 - 1300), color=RED, number=4, badge="bl"),
    ]
    annotate(cropB, boxesB)
    save(cropB, "asis_brand_B.png")

    # Crop C: 2140-2460 (申請フロー①〜⑥ ⑤B-3)
    cropC = im.crop((0, 2140, 1463, 2460))
    boxesC = [
        dict(xyxy=(173, 2317 - 2140, 1130, 2435 - 2140), color=RED, number=5, badge="tr"),
    ]
    annotate(cropC, boxesC)
    save(cropC, "asis_brand_C.png")

    # 縮小サムネイル(全体) + B-1 リーダー線(ページ全体を貫く)
    thumb_w = 210
    scale = thumb_w / im.width
    thumb = im.resize((thumb_w, int(im.height * scale)), Image.LANCZOS)
    draw = ImageDraw.Draw(thumb)
    x = thumb_w - 14
    # 破線を手描き
    y = 8
    while y < thumb.height - 8:
        draw.line([(x, y), (x, min(y + 14, thumb.height - 8))], fill=RED, width=5)
        y += 24
    draw_number_badge(thumb, draw, (x, 22), "B-1", RED, size=40)
    draw_number_badge(thumb, draw, (x, thumb.height - 22), "B-1", RED, size=40)
    save(thumb, "asis_brand_thumbnail.png")

    # B-7比較用: PC(8203px) vs SP(12196px) を同縮尺で並べる
    sp = open_src("brand_sp.png")
    px_per_orig = 0.045  # 縮尺
    pc_th = im.resize((int(im.width * px_per_orig), int(im.height * px_per_orig)), Image.LANCZOS)
    sp_th = sp.resize((int(sp.width * px_per_orig), int(sp.height * px_per_orig)), Image.LANCZOS)
    gap = 90
    top_margin, bottom_margin = 20, 50
    canvas_w = pc_th.width + sp_th.width + gap * 3
    canvas_h = max(pc_th.height, sp_th.height) + top_margin + bottom_margin
    canvas = Image.new("RGB", (canvas_w, canvas_h), WHITE)
    canvas.paste(pc_th, (gap, top_margin))
    canvas.paste(sp_th, (gap * 2 + pc_th.width, top_margin))
    d = ImageDraw.Draw(canvas)
    f = font(20)

    def centered_label(cx, y, text):
        bbox = d.textbbox((0, 0), text, font=f)
        tw = bbox[2] - bbox[0]
        d.text((cx - tw / 2 - bbox[0], y), text, fill=(20, 20, 20), font=f)

    centered_label(gap + pc_th.width / 2, top_margin + pc_th.height + 8, f"PC版 {im.height:,}px")
    centered_label(gap * 2 + pc_th.width + sp_th.width / 2, top_margin + sp_th.height + 8, f"SP版 {sp.height:,}px")
    d.rectangle([gap, top_margin, gap + pc_th.width, top_margin + pc_th.height], outline=RED, width=3)
    d.rectangle([gap * 2 + pc_th.width, top_margin, gap * 2 + pc_th.width + sp_th.width, top_margin + sp_th.height], outline=RED, width=3)
    save(canvas, "asis_brand_pc_vs_sp.png")


# ---------------------------------------------------------------------------
# 3. クライアント向け (client_pc)
# ---------------------------------------------------------------------------
def build_client():
    im = open_src("client_pc.png")

    crop1 = im.crop((0, 0, 1463, 1250))
    boxes1 = [dict(xyxy=(160, 415, 1020, 460), color=BLUE, good=True, number=1)]
    annotate(crop1, boxes1)
    save(crop1, "asis_client_1.png")

    crop2 = im.crop((0, 2200, 1463, 2900))
    boxes2 = [
        dict(xyxy=(170, 2263 - 2200, 1210, 2850 - 2200), color=RED, number=2, badge="tl"),
        dict(xyxy=(200, 2295 - 2200, 1180, 2333 - 2200), color=RED, number=3, badge="tr"),
    ]
    annotate(crop2, boxes2)
    save(crop2, "asis_client_2.png")


# ---------------------------------------------------------------------------
# 4. キャンディデート向け (candidate_pc) — 最重要
# ---------------------------------------------------------------------------
def build_candidate():
    im = open_src("candidate_pc.png")
    boxes = [
        dict(xyxy=(70, 385, 1390, 1795), color=BLUE, good=True, number=2),
        dict(xyxy=(160, 1085, 1210, 1230), color=RED, number=1),
    ]
    # Goodは全体を囲むので先に描画してからRedを上書きする(視認性のため順序調整)
    draw = ImageDraw.Draw(im)
    draw_box(draw, (70, 385, 1390, 1795), BLUE)
    draw_good_badge(im, draw, (76, 391))
    draw_box(draw, (160, 1085, 1210, 1230), RED)
    draw_number_badge(im, draw, (1210, 1085), 1, RED, size=44)
    save(im, "asis_candidate.png")


# ---------------------------------------------------------------------------
# 5. グローバルブランドパートナーシップ (partner_pc)
# ---------------------------------------------------------------------------
def build_partner():
    im = open_src("partner_pc.png")

    crop1 = im.crop((0, 2450, 1463, 3050))
    boxes1 = [dict(xyxy=(585, 2627 - 2450, 947, 2988 - 2450), color=RED, number=1)]
    annotate(crop1, boxes1)
    save(crop1, "asis_partner_1.png")

    thumb_w = 170
    scale = thumb_w / im.width
    thumb = im.resize((thumb_w, int(im.height * scale)), Image.LANCZOS)
    draw = ImageDraw.Draw(thumb)
    x = thumb_w - 12
    y = 8
    while y < thumb.height - 8:
        draw.line([(x, y), (x, min(y + 14, thumb.height - 8))], fill=RED, width=5)
        y += 24
    draw_number_badge(thumb, draw, (x, 22), 2, RED, size=36)
    draw_number_badge(thumb, draw, (x, thumb.height - 22), 2, RED, size=36)
    save(thumb, "asis_partner_thumbnail.png")


# ---------------------------------------------------------------------------
# 6. 横断課題 G-1〜G-4 (複数ページから代表crop)
# ---------------------------------------------------------------------------
def build_cross():
    client = open_src("client_pc.png")
    g1_client = client.crop((100, 380, 1080, 470))
    d = ImageDraw.Draw(g1_client)
    draw_box(d, (60, 35, 920, 80), BLUE)
    draw_good_badge(g1_client, d, (60, -6 if False else 2))
    save(g1_client, "asis_g1_client.png")

    brand = open_src("brand_pc.png")
    g1_brand = brand.crop((100, 370, 1080, 470))
    d = ImageDraw.Draw(g1_brand)
    draw_box(d, (60, 20, 920, 80), RED)
    draw_number_badge(g1_brand, d, (60, 20), "×", RED, size=40)
    save(g1_brand, "asis_g1_brand.png")

    top = open_src("top_pc.png")
    g2 = top.crop((560, 0, 970, 55))
    d = ImageDraw.Draw(g2)
    draw_box(d, (5, 5, 405, 50), RED)
    save(g2, "asis_g2_search.png")

    g3_top = top.crop((550, 880, 970, 970))
    d = ImageDraw.Draw(g3_top)
    draw_box(d, (10, 10, 390, 80), RED)
    save(g3_top, "asis_g3_top.png")

    g3_brand = brand.crop((600, 525, 970, 565))
    d = ImageDraw.Draw(g3_brand)
    draw_box(d, (5, 5, 365, 35), RED)
    save(g3_brand, "asis_g3_brand.png")

    candidate = open_src("candidate_pc.png")
    g3_candidate = candidate.crop((150, 1090, 1000, 1130))
    d = ImageDraw.Draw(g3_candidate)
    draw_box(d, (5, 5, 845, 35), RED)
    save(g3_candidate, "asis_g3_candidate.png")

    g4 = top.crop((1190, 108, 1375, 158))
    d = ImageDraw.Draw(g4)
    draw_box(d, (8, 8, 177, 42), RED)
    save(g4, "asis_g4_nav.png")


if __name__ == "__main__":
    build_top()
    build_brand()
    build_client()
    build_candidate()
    build_partner()
    build_cross()
    print("done")

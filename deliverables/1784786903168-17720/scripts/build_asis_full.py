#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AsIs フルページ注釈画像 生成スクリプト（Excel提案書 用・判読性優先）
- materials/screenshots/ の実スクショ(PC版・全長)をそのまま使用（トリミングしない）
- 課題ID付き番号バッジ（赤=課題／青=Good=参照価値）を実座標に描画
- 画像下部に凡例（番号→課題ID→一言説明）を追加し、画像単体で意味が分かるようにする
- 出力: deliverables/1784786903168-17720/assets/asis_full_{top,brand,client,candidate,partner}.png
"""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = "/workspace/akkodis-marketing-portal"
SRC = os.path.join(ROOT, "materials/screenshots")
ASSETS = os.path.join(ROOT, "deliverables/1784786903168-17720/assets")
OUT = ASSETS
os.makedirs(OUT, exist_ok=True)

RED = (196, 30, 30)
BLUE = (0, 90, 181)
WHITE = (255, 255, 255)
DARK = (25, 25, 25)
LEGEND_BG = (247, 247, 245)

FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"


def font(size):
    return ImageFont.truetype(FONT_PATH, size)


def open_src(name):
    return Image.open(os.path.join(SRC, name)).convert("RGB")


def draw_box(draw, xyxy, color, halo_w=8, line_w=5):
    x0, y0, x1, y1 = xyxy
    draw.rectangle([x0 - halo_w // 2, y0 - halo_w // 2, x1 + halo_w // 2, y1 + halo_w // 2],
                   outline=WHITE, width=halo_w)
    draw.rectangle([x0, y0, x1, y1], outline=color, width=line_w)


def draw_badge(draw, center, text, color, size=48):
    """課題ID（T-1／CA-2 等・2〜4文字）を丸ピル型バッジで描画。文字幅に応じて自動で幅を広げ、
    テキストがバッジ枠からはみ出さないようにする（判読性優先）。"""
    cx, cy = center
    f = font(int(size * 0.42))
    bbox = draw.textbbox((0, 0), text, font=f)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pad_x, pad_y = 16, 10
    h = max(size, th + pad_y * 2)
    w = max(h, tw + pad_x * 2)
    x0, y0, x1, y1 = cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2
    draw.rounded_rectangle([x0, y0, x1, y1], radius=h / 2, fill=color, outline=WHITE, width=5)
    draw.text((cx - tw / 2 - bbox[0], cy - th / 2 - bbox[1]), text, fill=WHITE, font=f)


def dashed_vline(draw, x, y0, y1, color, dash=16, gap=10, width=6):
    y = y0
    while y < y1:
        y2 = min(y + dash, y1)
        draw.line([(x, y), (x, y2)], fill=color, width=width)
        y = y2 + gap


def annotate(im, items):
    """items: list of dict(xyxy, number(str), good=False)"""
    draw = ImageDraw.Draw(im)
    for it in items:
        color = BLUE if it.get("good") else RED
        draw_box(draw, it["xyxy"], color, halo_w=10, line_w=6)
        x0, y0, x1, y1 = it["xyxy"]
        pos = it.get("badge", "tl")
        if pos == "tr":
            bx, by = x1, y0
        elif pos == "bl":
            bx, by = x0, y1
        else:
            bx, by = x0, y0
        draw_badge(draw, (bx, by), it["number"], color, size=56)
    return im


def add_legend(im, rows, title, page_w):
    """rows: list of (number, color, text). Appends a legend strip below the image."""
    pad = 24
    line_h = 40
    f_title = font(30)
    f_row = font(24)
    header_h = 56
    height = header_h + pad * 2 + line_h * len(rows)
    legend = Image.new("RGB", (page_w, height), LEGEND_BG)
    d = ImageDraw.Draw(legend)
    d.rectangle([0, 0, page_w - 1, height - 1], outline=(200, 200, 195), width=2)
    d.text((pad, pad - 6), title, fill=DARK, font=f_title)
    y = header_h + pad
    for number, color, text in rows:
        cx, cy = pad + 34, y + line_h // 2 - 6
        f2 = font(20)
        bbox = d.textbbox((0, 0), number, font=f2)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        pad_x, pad_y = 12, 7
        h = max(36, th + pad_y * 2)
        w = max(h, tw + pad_x * 2)
        d.rounded_rectangle([cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2], radius=h / 2, fill=color, outline=WHITE, width=3)
        d.text((cx - tw / 2 - bbox[0], cy - th / 2 - bbox[1]), number, fill=WHITE, font=f2)
        d.text((pad + 76, y), text, fill=DARK, font=f_row)
        y += line_h

    canvas = Image.new("RGB", (page_w, im.height + height), WHITE)
    canvas.paste(im, (0, 0))
    canvas.paste(legend, (0, im.height))
    return canvas


def top_banner(im, text, color=RED, h=64):
    """画像上部にバナー注記を追加（ページ全体に関わる注記用）"""
    w = im.width
    banner = Image.new("RGB", (w, h), (255, 244, 224))
    d = ImageDraw.Draw(banner)
    d.rectangle([0, 0, w - 1, h - 1], outline=color, width=3)
    f = font(26)
    d.text((20, h // 2 - 16), text, fill=DARK, font=f)
    canvas = Image.new("RGB", (w, im.height + h), WHITE)
    canvas.paste(banner, (0, 0))
    canvas.paste(im, (0, h))
    return canvas


def save(im, name):
    path = os.path.join(OUT, name)
    im.save(path, optimize=True)
    print("saved", path, im.size, round(os.path.getsize(path) / 1024 / 1024, 2), "MB")


# ---------------------------------------------------------------------------
# 1. トップ
# ---------------------------------------------------------------------------
def build_top():
    im = open_src("top_pc.png")
    items = [
        dict(xyxy=(160, 410, 1375, 626), number="T-3"),
        dict(xyxy=(568, 898, 954, 953), number="T-4", badge="tr"),
        dict(xyxy=(160, 1012, 780, 1046), number="T-4", badge="bl"),
        dict(xyxy=(80, 2630, 1385, 3200), number="T-2"),
    ]
    annotate(im, items)

    # T-1（SP限定の課題）はPC画面に写らないため、SP版の該当箇所を小さく切り出してインセット表示
    sp = open_src("top_sp.png")
    inset = sp.crop((0, 2080, 458, 2720))
    d = ImageDraw.Draw(inset)
    draw_box(d, (15, 170, 440, 615), RED)
    draw_badge(d, (440, 170), "T-1", RED, size=48)
    # 枠線とキャプションを付けてインセットに仕立てる
    cap_h = 40
    box = Image.new("RGB", (inset.width + 8, inset.height + cap_h + 8), WHITE)
    dd = ImageDraw.Draw(box)
    dd.rectangle([0, 0, box.width - 1, box.height - 1], outline=RED, width=4)
    box.paste(inset, (4, 4))
    dd.text((10, inset.height + 8), "SP版：T-1 埋め込みダッシュボード極小表示", fill=DARK, font=font(20))

    canvas = Image.new("RGB", (im.width, im.height), WHITE)
    canvas.paste(im, (0, 0))
    canvas.paste(box, (im.width - box.width - 20, 20))
    im = canvas

    rows = [
        ("T-1", RED, "埋め込みダッシュボードがモバイル非対応（SP版で極小表示・右上インセット参照）"),
        ("T-2", RED, "「組織体制」が最下部に配置され、探す・依頼する導線より優先順位が低く見える"),
        ("T-3", RED, "4カテゴリカードがナビと重複し、説明文がなく「次に何ができるか」が伝わらない"),
        ("T-4", RED, "業務依頼フォームとブランドレビュー手順など複数入口の使い分けが1画面で判別しづらい"),
    ]
    im = add_legend(im, rows, "AsIs（現状）注釈凡例 — トップ（ホーム）", im.width)
    save(im, "asis_full_top.png")


# ---------------------------------------------------------------------------
# 2. ブランドライブラリ（最重要・全長8203px）
# ---------------------------------------------------------------------------
def build_brand():
    im = open_src("brand_pc.png")
    items = [
        dict(xyxy=(160, 465, 1300, 945), number="B-6"),
        dict(xyxy=(155, 1108, 1310, 1310), number="B-5", badge="bl"),
        dict(xyxy=(185, 1558, 1360, 1795), number="B-2", badge="tr"),
        dict(xyxy=(190, 1558, 700, 1598), number="B-4", badge="bl"),
        dict(xyxy=(173, 2317, 1130, 2435), number="B-3", badge="tr"),
    ]
    annotate(im, items)

    draw = ImageDraw.Draw(im)
    # B-1: 情報過多・ページ内ナビ欠如＝ページ全体にまたがる構造課題。左端の赤破線で「1本の長尺ページ」を可視化
    dashed_vline(draw, 24, 40, im.height - 40, RED, dash=26, gap=16, width=8)
    draw_badge(draw, (24, 60), "B-1", RED, size=56)
    draw_badge(draw, (24, im.height - 60), "B-1", RED, size=56)

    im = top_banner(
        im,
        f"B-7: SP版はさらに長尺（PC {im.height - 64:,}px超 → SP 12,196px相当）。モバイルでの回遊が困難",
    )

    rows = [
        ("B-1", RED, "情報過多・深いネスト・ページ内ナビ（目次/アンカー）欠如（左端破線＝1ページの全長）"),
        ("B-2", RED, "ガイドライン約10本がフラットに列挙され、目的別に束ねられていない"),
        ("B-3", RED, "ロゴ掲載フローがフォーム→Teams→Salesforceをまたぐ6ステップを長文で説明"),
        ("B-4", RED, "資料の命名規則が不統一（Akkodis/AKKODiS表記ゆれ・版/言語/形式の混在）"),
        ("B-5", RED, "ラベルのないアイコンカード3枚。アフォーダンス不明"),
        ("B-6", RED, "冒頭見出しの黄色マーカーが不統一（帯見出しとの二重表現）"),
        ("B-7", RED, "上部バナー参照：SP版が極端に長尺（1万px超）でモバイル回遊が困難"),
    ]
    im = add_legend(im, rows, "AsIs（現状）注釈凡例 — ブランドライブラリ（最重要）", im.width)
    save(im, "asis_full_brand.png")


# ---------------------------------------------------------------------------
# 3. クライアント向け
# ---------------------------------------------------------------------------
def build_client():
    im = open_src("client_pc.png")
    items = [
        dict(xyxy=(160, 415, 1020, 460), number="C-3", good=True),
        dict(xyxy=(170, 2263, 1210, 2850), number="C-1", badge="tl"),
        dict(xyxy=(200, 2295, 1180, 2333), number="C-2", badge="tr"),
    ]
    annotate(im, items)
    rows = [
        ("C-3", BLUE, "Good：冒頭アンカー型クイックリンクが機能（他ページへの横展開を推奨）"),
        ("C-1", RED, "「お役立ち資料」が30本超のリンク羅列。説明文・サムネ・絞り込みがない"),
        ("C-2", RED, "リンクに更新日・対象・要約などのメタ情報がなく最新版/対象者が判別できない"),
    ]
    im = add_legend(im, rows, "AsIs（現状）注釈凡例 — クライアント向け", im.width)
    save(im, "asis_full_client.png")


# ---------------------------------------------------------------------------
# 4. キャンディデート向け
# ---------------------------------------------------------------------------
def build_candidate():
    im = open_src("candidate_pc.png")
    draw = ImageDraw.Draw(im)
    draw_box(draw, (70, 385, 1390, 1795), BLUE, halo_w=10, line_w=6)
    draw_badge(draw, (70, 385), "CA-2", BLUE, size=56)
    draw_box(draw, (160, 1085, 1210, 1230), RED, halo_w=10, line_w=6)
    draw_badge(draw, (1210, 1085), "CA-1", RED, size=56)
    rows = [
        ("CA-2", BLUE, "Good：5ページ中もっとも簡潔で見やすい構成（他ページの目標形にできる参照モデル）"),
        ("CA-1", RED, "マスター資料が「各チームに直接連絡」で完結し、資料そのものが入手できない"),
    ]
    im = add_legend(im, rows, "AsIs（現状）注釈凡例 — キャンディデート向け", im.width)
    save(im, "asis_full_candidate.png")


# ---------------------------------------------------------------------------
# 5. グローバルブランドパートナーシップ
# ---------------------------------------------------------------------------
def build_partner():
    im = open_src("partner_pc.png")
    items = [
        dict(xyxy=(585, 2627, 947, 2988), number="P-1"),
    ]
    annotate(im, items)
    draw = ImageDraw.Draw(im)
    # P-2: ページ内ナビ欠如＝ページ全体の構造課題（右端の赤破線で全長を可視化）
    dashed_vline(draw, im.width - 24, 40, im.height - 40, RED, dash=26, gap=16, width=8)
    draw_badge(draw, (im.width - 24, 60), "P-2", RED, size=56)
    draw_badge(draw, (im.width - 24, im.height - 60), "P-2", RED, size=56)
    rows = [
        ("P-1", RED, "Akkodis ASP Teamのみ余分な（壊れた）画像枠がプレースホルダー表示（要現物確認）"),
        ("P-2", RED, "長尺だがページ内ナビがない（右端破線＝1ページの全長。クライアントページと不統一）"),
    ]
    im = add_legend(im, rows, "AsIs（現状）注釈凡例 — グローバルブランドパートナーシップ", im.width)
    save(im, "asis_full_partner.png")


if __name__ == "__main__":
    build_top()
    build_brand()
    build_client()
    build_candidate()
    build_partner()
    print("done")

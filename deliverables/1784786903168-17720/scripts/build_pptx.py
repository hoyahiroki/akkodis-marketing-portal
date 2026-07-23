#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AKKODiS マーケティングポータル刷新提案書 pptx 組版スクリプト
- 16:9 / 左AsIs・右ToBe対比 / 全スライド社外秘フッター
- python-pptx で assets/ 配下の生成画像を配置し、凡例・改善ポイントはネイティブテキストで付与
"""
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ROOT = "/workspace/akkodis-marketing-portal"
ASSETS = os.path.join(ROOT, "deliverables/1784786903168-17720/assets")
OUT_PATH = os.path.join(ROOT, "deliverables/1784786903168-17720/AKKODiSマーケティングポータル_課題AsIs-ToBe提案書.pptx")
LOGO = "/root/.claude/skills/akkodis-brand-core/assets/logos/AKKODIS_Logo_POS_RGB.png"

NAVY = RGBColor(0x00, 0x1F, 0x33)
YELLOW = RGBColor(0xFF, 0xB8, 0x1C)
CYAN = RGBColor(0x00, 0xB8, 0xC4)  # 印刷可読性のため本家シアンより少し沈めたトーン(本文使用ではないため逸脱可)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARKGRAY = RGBColor(0x33, 0x33, 0x33)
MIDGRAY = RGBColor(0x66, 0x66, 0x66)
RED = RGBColor(0xC4, 0x1E, 0x1E)
BLUE = RGBColor(0x00, 0x5A, 0xB5)
LIGHTBG = RGBColor(0xF4, 0xF5, 0xF6)

FONT_JP = "Meiryo UI"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]

PAGE_NO = [0]


def new_slide():
    return prs.slides.add_slide(BLANK)


def set_font(run, size=11, bold=False, color=DARKGRAY, italic=False, font=FONT_JP):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    ea = rPr.makeelement(qn('a:ea'), {'typeface': font})
    rPr.append(ea)


def add_rect(slide, left, top, width, height, fill=None, line_color=None, line_w=1.0, shadow=False, dashed=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(line_w)
        if dashed:
            d = shp.line._get_or_add_ln()
            pd = d.makeelement(qn('a:prstDash'), {'val': 'dash'})
            d.append(pd)
    return shp


def add_text(slide, left, top, width, height, text, size=11, bold=False, color=DARKGRAY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=FONT_JP, line_spacing=1.15,
             wrap=True, shrink=False, italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        set_font(r, size=size, bold=bold, color=color, font=font, italic=italic)
    return tb


def add_multirun_bullet(slide, left, top, width, height, items, size=10.5, color=DARKGRAY,
                         bullet_color=NAVY, gap_pt=4, font=FONT_JP):
    """items: list of str. 各行冒頭に ▪ を付ける簡易箇条書き"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.12
        p.space_after = Pt(gap_pt)
        r = p.add_run()
        r.text = "・ " + item
        set_font(r, size=size, bold=False, color=color, font=font)
    return tb


def img_size(path):
    with Image.open(path) as im:
        return im.size


def add_picture_fit(slide, path, left, top, max_w, max_h, valign="top"):
    iw, ih = img_size(path)
    ratio = min(max_w / iw, max_h / ih)
    w = int(iw * ratio)
    h = int(ih * ratio)
    l = left + (max_w - w) // 2
    if valign == "top":
        t = top
    elif valign == "middle":
        t = top + (max_h - h) // 2
    else:
        t = top + (max_h - h)
    slide.shapes.add_picture(path, l, t, width=w, height=h)
    return l, t, w, h


def footer(slide, title_short):
    add_rect(slide, 0, Inches(7.2), SLIDE_W, Inches(0.3), fill=RGBColor(0xEC, 0xEE, 0xF0))
    add_text(slide, Inches(0.3), Inches(7.21), Inches(9), Inches(0.28),
              f"社外秘（Confidential）— AKKODiSコンサルティング株式会社 マーケティング部内限り｜{title_short}",
              size=8.5, color=MIDGRAY)
    PAGE_NO[0] += 1
    add_text(slide, Inches(12.6), Inches(7.21), Inches(0.6), Inches(0.28), str(PAGE_NO[0]),
              size=8.5, color=MIDGRAY, align=PP_ALIGN.RIGHT)


def title_bar(slide, title, subtitle=None, badge=None):
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.85), fill=NAVY)
    add_rect(slide, 0, Inches(0.85), SLIDE_W, Pt(3), fill=YELLOW)
    add_text(slide, Inches(0.35), Inches(0.1), Inches(10.5), Inches(0.5), title,
              size=22, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, Inches(0.35), Inches(0.55), Inches(10.5), Inches(0.3), subtitle,
                  size=11, color=RGBColor(0xCF, 0xE3, 0xF0))
    if badge:
        w = Inches(2.6)
        add_rect(slide, SLIDE_W - w - Inches(0.3), Inches(0.18), w, Inches(0.5), fill=YELLOW)
        add_text(slide, SLIDE_W - w - Inches(0.3), Inches(0.18), w, Inches(0.5), badge,
                  size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def column_label(slide, left, text, color):
    w = Inches(2.3)
    add_rect(slide, left, Inches(0.98), w, Inches(0.32), fill=color)
    add_text(slide, left, Inches(0.98), w, Inches(0.32), text, size=12.5, bold=True,
              color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


LEFT_X = Inches(0.35)
LEFT_W = Inches(6.15)
RIGHT_X = Inches(6.75)
RIGHT_W = Inches(6.25)
COL_TOP = Inches(1.40)
POINTS_BOTTOM = Inches(7.12)


def render_row_images(slide, x, y, w, h, images):
    """images: list of (path, caption). 同一行内に横並びでアスペクト比を保ち収める(幅超過時は縮小)"""
    if not images:
        return
    has_caption = any(cap for _, cap in images)
    cap_h = Inches(0.16) if has_caption else 0
    avail_h = h - cap_h
    gap = Inches(0.1)
    sizes = []
    for path, cap in images:
        iw, ih = img_size(path)
        ww = avail_h * iw / ih
        sizes.append((path, cap, ww))
    total_w = sum(s[2] for s in sizes) + gap * (len(sizes) - 1)
    scale = min(1.0, w / total_w) if total_w > 0 else 1.0
    row_w = total_w * scale
    cx = x + (w - row_w) / 2
    for path, cap, ww in sizes:
        wpx = int(ww * scale)
        hpx = int(avail_h * scale)
        slide.shapes.add_picture(path, int(cx), int(y), width=wpx, height=hpx)
        if cap:
            capw = max(wpx, Inches(1.3))
            capx = cx + wpx / 2 - capw / 2
            add_text(slide, int(capx), int(y + hpx + Inches(0.01)), int(capw), Inches(0.15), cap,
                      size=7.5, color=MIDGRAY, align=PP_ALIGN.CENTER)
        cx += wpx + gap


def render_rows(slide, x, y, w, h, rows):
    """rows: list of (images, weight)"""
    total_weight = sum(wt for _, wt in rows) or 1
    cy = y
    for images, wt in rows:
        row_h = int(h * wt / total_weight)
        render_row_images(slide, x, cy, w, row_h, images)
        cy += row_h


def legend_block(slide, left, top, width, items, title="AsIs 該当箇所"):
    add_text(slide, left, top, width, Inches(0.22), title, size=10, bold=True, color=NAVY)
    y = top + Inches(0.26)
    tb = slide.shapes.add_textbox(left, y, width, Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_top = 0
    for i, (num, good, text) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.08
        p.space_after = Pt(2)
        r1 = p.add_run()
        r1.text = f"{num} "
        set_font(r1, size=9.5, bold=True, color=(BLUE if good else RED))
        r2 = p.add_run()
        r2.text = text
        set_font(r2, size=9.5, color=DARKGRAY)
    return tb


def points_block(slide, items, height=Inches(1.15), size=9.3):
    top = POINTS_BOTTOM - height
    add_rect(slide, Inches(0.3), top, SLIDE_W - Inches(0.6), height, fill=LIGHTBG, line_color=RGBColor(0xC7, 0xCC, 0xD1), line_w=0.75)
    add_text(slide, Inches(0.45), top + Inches(0.05), Inches(4.0), Inches(0.22), "改善ポイント（課題→改善→効果）",
              size=10.5, bold=True, color=NAVY)
    add_multirun_bullet(slide, Inches(0.45), top + Inches(0.3), SLIDE_W - Inches(0.9), height - Inches(0.36), items, size=size, gap_pt=3)
    return top


# ===========================================================================
# Slide 1: タイトル
# ===========================================================================
def slide_title():
    s = new_slide()
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, fill=NAVY)
    add_rect(s, 0, Inches(4.55), SLIDE_W, Pt(4), fill=YELLOW)
    try:
        iw, ih = img_size(LOGO)
        w = Inches(2.6)
        h = int(w * ih / iw)
        s.shapes.add_picture(LOGO, Inches(0.9), Inches(0.75), width=w, height=h)
    except Exception:
        pass
    add_text(s, Inches(0.9), Inches(2.15), Inches(11.6), Inches(1.6),
              "AKKODiSマーケティングポータル\nリニューアル提案", size=30, bold=True, color=WHITE, line_spacing=1.15)
    add_text(s, Inches(0.9), Inches(3.85), Inches(11.0), Inches(0.5),
              "現状課題22件の分析とSharePoint L1既存パーツによる改善方向性", size=16, color=RGBColor(0xCF, 0xE3, 0xF0))
    add_text(s, Inches(0.9), Inches(4.85), Inches(9), Inches(0.4),
              "AsIs / ToBe 対比資料　―　デザイン部署", size=12, color=RGBColor(0xB8, 0xC7, 0xD1))
    add_text(s, Inches(0.9), Inches(6.7), Inches(6), Inches(0.4),
              "社外秘（Confidential）— AKKODiSコンサルティング株式会社 マーケティング部内限り", size=10.5, color=RGBColor(0x8C, 0xA0, 0xAC))
    PAGE_NO[0] += 1
    add_text(s, Inches(12.6), Inches(7.05), Inches(0.6), Inches(0.3), str(PAGE_NO[0]), size=9, color=RGBColor(0x8C,0xA0,0xAC), align=PP_ALIGN.RIGHT)


# ===========================================================================
# Slide 2: エグゼクティブサマリー
# ===========================================================================
def slide_exec_summary():
    s = new_slide()
    title_bar(s, "エグゼクティブサマリー", "現状分析（docs/01）とSharePoint実装マッピング（docs/02）の要点")

    # 課題内訳
    add_text(s, Inches(0.35), Inches(1.05), Inches(6.0), Inches(0.26), "現状課題 22件の内訳", size=13, bold=True, color=NAVY)
    counts = [("トップ", "T-1〜T-4", "4件"), ("ブランドライブラリ", "B-1〜B-7", "7件"),
              ("クライアント向け", "C-1〜C-3", "3件"), ("キャンディデート向け", "CA-1〜CA-2", "2件"),
              ("パートナーシップ", "P-1〜P-2", "2件"), ("横断課題", "G-1〜G-4", "4件")]
    cx = Inches(0.35)
    cw = Inches(1.98)
    for i, (name, rng, n) in enumerate(counts):
        cxi = cx + Emu(int(cw) * (i % 3))
        cyi = Inches(1.38) + Emu(int(Inches(0.72)) * (i // 3))
        add_rect(s, cxi, cyi, cw - Inches(0.08), Inches(0.64), fill=LIGHTBG, line_color=RGBColor(0xC7,0xCC,0xD1), line_w=0.75)
        add_text(s, cxi + Inches(0.08), cyi + Inches(0.05), cw - Inches(0.2), Inches(0.24), f"{name}", size=10, bold=True, color=NAVY)
        add_text(s, cxi + Inches(0.08), cyi + Inches(0.28), cw - Inches(0.2), Inches(0.18), rng, size=8, color=MIDGRAY)
        add_text(s, cxi + Inches(0.08), cyi + Inches(0.28), cw - Inches(0.24), Inches(0.32), n, size=14, bold=True, color=RED, align=PP_ALIGN.RIGHT)

    # 最重要3課題
    top_y = Inches(2.95)
    add_text(s, Inches(0.35), top_y, Inches(6.0), Inches(0.26), "最重要3課題（優先着手）", size=13, bold=True, color=NAVY)
    top3 = [
        ("B-1", "ブランドライブラリの情報過多・ページ内ナビ欠如", "直近30日で最も使われる実働ページ（重複しない閲覧者数319名）かつユーザー指摘に完全一致。改善効果が最も広く届く。"),
        ("G-2", "ポータル内検索性の弱さ", "「見つからない→メール検索」という核心症状の根本原因。標準検索はあるがページ内の絞り込み・ファセットが無い。"),
        ("CA-1", "資料が「各チームに直接連絡」で完結し入手不可", "ポータルがメール検索を誘発している典型例。直リンク／保管場所の明示で即改善可能なクイックウィン。"),
    ]
    ty = top_y + Inches(0.3)
    for tag, name, desc in top3:
        add_rect(s, Inches(0.35), ty, Inches(0.62), Inches(0.5), fill=RED)
        add_text(s, Inches(0.35), ty, Inches(0.62), Inches(0.5), tag, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(1.08), ty, Inches(5.35), Inches(0.24), name, size=10.5, bold=True, color=DARKGRAY)
        add_text(s, Inches(1.08), ty + Inches(0.24), Inches(5.35), Inches(0.55), desc, size=8.8, color=MIDGRAY, line_spacing=1.05)
        ty += Inches(0.72)

    # 右カラム: 改善アプローチ + MVP4点
    rx = Inches(6.9)
    add_text(s, rx, Inches(1.05), Inches(6.1), Inches(0.26), "改善アプローチ：方針転換", size=13, bold=True, color=NAVY)
    add_rect(s, rx, Inches(1.36), Inches(6.05), Inches(1.0), fill=LIGHTBG, line_color=RGBColor(0xC7,0xCC,0xD1), line_w=0.75)
    add_text(s, rx + Inches(0.15), Inches(1.45), Inches(5.75), Inches(0.85),
              "初版は「AI機能ありき」で設計していたが、社長差戻し（SharePoint実装マスト）を受け優先順位を反転。"
              "「SharePoint既存パーツ（L1）で今すぐ作れるもの」を起点に据え、それを超える機能（AI-1〜AI-6等）は"
              "実装階層L1〜L4を明示してライセンス・費用判断に切り出す。",
              size=9.3, color=DARKGRAY, line_spacing=1.15)

    add_text(s, rx, Inches(2.55), Inches(6.1), Inches(0.26), "MVP（L1・4点）", size=13, bold=True, color=NAVY)
    mvp = [
        "キャンディデート資料の直リンク化（CA-1／クイックウィン）",
        "ブランドライブラリ再構成＋ページ内ナビ（B-1／B-2）",
        "Microsoft Search 基本＋検索・ハイライトweb part設置（G-2）",
        "ライブラリのメタデータ列・コンテンツタイプ設計（B-4／C-2の土台）",
    ]
    add_multirun_bullet(s, rx, Inches(2.85), Inches(6.05), Inches(1.6), mvp, size=10, gap_pt=8)

    add_rect(s, rx, Inches(4.55), Inches(6.05), Inches(2.3), fill=RGBColor(0xFB, 0xF3, 0xE8), line_color=RGBColor(0xB5,0x76,0x2A), line_w=1.0, dashed=True)
    add_text(s, rx + Inches(0.15), Inches(4.65), Inches(5.7), Inches(0.24), "本書の位置づけ", size=10.5, bold=True, color=RGBColor(0x8a,0x55,0x18))
    add_text(s, rx + Inches(0.15), Inches(4.95), Inches(5.75), Inches(1.8),
              "本資料はデザイン部署によるAsIs→ToBe対比の可視化提案であり、実装・費用発生・外部公開の意思決定は含まない。"
              "SharePointの機能名・パーツ名は知識カットオフ時点の一般知識に基づく推定を含み、テナントの実ライセンス・"
              "管理者ポリシーでの検証（F0）を前提とする。文言未確定箇所（カード説明文・アイコンラベル等）は"
              "「（今後定義）」と明記し、創作した具体文言を確定情報のようには見せていない。",
              size=8.8, color=DARKGRAY, line_spacing=1.15)

    footer(s, "エグゼクティブサマリー")


# ===========================================================================
# 汎用: AsIs/ToBe対比スライド
# ===========================================================================
def compare_slide(title, subtitle, badge, asis_rows, tobe_images, legend_items, points, legend_title="AsIs 該当箇所",
                   asis_col_title="AsIs（現状）", tobe_col_title="ToBe（改善後）", points_height=Inches(1.05),
                   legend_height=Inches(0.95)):
    s = new_slide()
    title_bar(s, title, subtitle, badge=badge)
    column_label(s, LEFT_X, asis_col_title, RED)
    column_label(s, RIGHT_X, tobe_col_title, BLUE)

    img_top = COL_TOP + Inches(0.40)
    points_top = points_block(s, points, height=points_height)
    content_bottom = points_top - Inches(0.1)

    left_img_h = content_bottom - img_top - legend_height - Inches(0.08)
    render_rows(s, LEFT_X, img_top, LEFT_W, left_img_h, asis_rows)

    legend_block(s, LEFT_X, img_top + left_img_h + Inches(0.1), LEFT_W, legend_items, title=legend_title)

    tobe_h = content_bottom - img_top
    render_rows(s, RIGHT_X, img_top, RIGHT_W, tobe_h, [(tobe_images, 1)])

    footer(s, title)


# ===========================================================================
# Slide 3: トップ
# ===========================================================================
def slide_top():
    legend = [
        ("①", False, "T-3　説明文がなくラベルのみで内容が伝わらない"),
        ("②", False, "T-4　似た入口が並び使い分けが1画面で判別できない"),
        ("③", False, "T-2　探す・依頼する導線より下に組織図が沈む"),
        ("④", False, "T-1（SP）　モバイルでは極小枠＋横スクロールで数値が読めない"),
    ]
    points = [
        "T-3：カードに説明文がなく次に何ができるか伝わらない → クイックリンクに1行説明を追加 → クリック前に目的地が分かり誤クリック・迷いが減る",
        "T-4：業務依頼フォームとブランドレビュー手順の使い分けが1画面で判別しづらい → 「よく使う窓口」セクションで2入口を並置し使い分けテキストを添える → 依頼者が正しい窓口へ一発到達",
        "T-1：埋め込みダッシュボードがSPで極小枠＋横スクロールとなり判読不能 → SP表示のみ要約表示に切替 → モバイル利用者にも指標が届く",
        "T-2：組織体制が最下部に沈み情報の優先順位が弱い → ページ内目次にアンカーを追加 → スクロールなしで直接到達可能に",
    ]
    compare_slide(
        "① トップ（ホーム）", "対象課題：T-1／T-2／T-3／T-4", "top_pc / top_sp",
        [([(f"{ASSETS}/asis_top_pc.png", "top_pc.png（①②③）"),
           (f"{ASSETS}/asis_top_sp_dashboard.png", "top_sp.png 抜粋（④）")], 1)],
        [(f"{ASSETS}/tobe_top.png", "")],
        legend, points, points_height=Inches(1.35),
    )


# ===========================================================================
# Slide 4-5: ブランドライブラリ（最重要・2スライド）
# ===========================================================================
def slide_brand_1():
    legend = [
        ("B-1", False, "（最重要）1ページ8000px超に全情報を積層、目次が皆無"),
        ("①", False, "B-6　帯見出しと黄色マーカーの装飾が二重で不統一"),
        ("②", False, "B-5　ラベルがなく何をするカードか分からない"),
    ]
    points = [
        "B-1（最重要）：1ページに全情報を積層しページ内ナビが皆無、目的の情報に辿り着けない → 7項目のページ内目次を冒頭に新設 → 直近30日で最多利用（319名）のページの到達性が改善し効果が最も広く届く",
        "B-2（最重要）：ガイドライン約10本が粒度・目的バラバラにフラット列挙 → 目的軸4カテゴリのドキュメントライブラリ＋メタデータ列で束ね直す → 目的の資料への選択時間が短縮",
    ]
    compare_slide(
        "② ブランドライブラリ（1/2）", "対象課題：B-1（最重要）／B-2（最重要）／B-6／B-5　※直近30日 最多閲覧ページ（319名）", "brand_pc（最重要）",
        [([(f"{ASSETS}/asis_brand_thumbnail.png", "全体縮尺（B-1）"),
           (f"{ASSETS}/asis_brand_A.png", "Crop A（①②）")], 1)],
        [(f"{ASSETS}/tobe_brand_1.png", "")],
        legend, points, points_height=Inches(1.2),
    )


def slide_brand_2():
    legend = [
        ("③", False, "B-2　粒度も種類も違う資料が10本フラットに並ぶ"),
        ("④", False, "B-4　Akkodis／AKKODiSなど表記が資料ごとに混在"),
        ("⑤", False, "B-3　フォーム→Teams→Salesforceを跨ぐ6手順が長文"),
        ("B-7", False, "SP版はPC版よりさらに長尺（12,196px／PC 8,203px）"),
    ]
    points = [
        "B-5：ラベルのない3枚のアイコンカードでアフォーダンスが不明 → クイックリンクにラベルテキストを追加 → クリック前に用途が分かる",
        "B-6：冒頭見出しの黄色マーカーが帯見出しと二重表現で体裁が粗い → 帯見出し1種類に統一 → 視覚的な一貫性が向上",
        "B-3：ロゴ掲載フローがフォーム→Teams→Salesforceを跨ぐ6ステップを長文で説明 → 番号付き手順＋関連テンプレ直リンクに整理 → 手続きの実行負荷を軽減（自動化は将来のL2拡張）",
        "B-7：モバイルは各セクションを折りたたみ表示にし、目次からのアンカー遷移を主動線にすることで長尺化を体感上緩和",
    ]
    compare_slide(
        "② ブランドライブラリ（2/2）", "対象課題：B-2／B-3／B-4／B-5／B-6／B-7", "brand_pc（最重要）",
        [([(f"{ASSETS}/asis_brand_B.png", "Crop B（③④）")], 0.34),
         ([(f"{ASSETS}/asis_brand_C.png", "Crop C（⑤）")], 0.24),
         ([(f"{ASSETS}/asis_brand_pc_vs_sp.png", "PC/SP 全長比較（B-7）")], 0.42)],
        [(f"{ASSETS}/tobe_brand_2.png", "")],
        legend, points, points_height=Inches(1.4),
    )


# ===========================================================================
# Slide 6: クライアント向け
# ===========================================================================
def slide_client():
    legend = [
        ("①", True, "C-3（Good）ページ内ジャンプが機能する唯一の良い例"),
        ("②", False, "C-1　30本超がメタ情報なしに羅列されるだけ"),
        ("③", False, "C-2　更新日・対象がなく新旧・要否が判別不可"),
    ]
    points = [
        "C-1：お役立ち資料が30本超のリンク羅列で説明文・絞り込みがない → メタデータ列付きドキュメントライブラリ＋カテゴリ別絞り込みに変更 → 目的資料の発見コストを削減",
        "C-2：リンクに更新日・対象などのメタ情報がなく最新版か判別できない → 言語・更新日メタデータ列を付与 → 新旧・要否の判別が可能に",
        "C-3：アンカーリンクが唯一機能している良い部分 → 構成をそのまま維持し他ページ（ブランド／パートナー）へ横展開 → 全ページで一貫した探索体験（G-1是正にも直結）",
    ]
    compare_slide(
        "③ クライアント向け", "対象課題：C-1／C-2／C-3（Good・参照モデル）", "client_pc",
        [([(f"{ASSETS}/asis_client_1.png", "Crop1（①）")], 0.48),
         ([(f"{ASSETS}/asis_client_2.png", "Crop2（②③）")], 0.52)],
        [(f"{ASSETS}/tobe_client.png", "")],
        legend, points, points_height=Inches(1.2),
    )


# ===========================================================================
# Slide 7: キャンディデート向け（最重要・厚く）
# ===========================================================================
def slide_candidate():
    legend = [
        ("①", False, "CA-1（最重要）直リンクがなく連絡しないと資料が手に入らない"),
        ("②", True, "CA-2（Good）3セクションのみで簡潔・他ページの参照モデル"),
    ]
    points = [
        "CA-1（最重要）：マスター資料が「各チームに直接連絡」で完結し資料が入手できない → ドキュメントライブラリへの直リンク＋クイックリンクに置換、連絡先はメタデータとして保持 → 「見つからずメール検索」を誘発する典型例を最短コストで即時解消（クイックウィン）",
        "CA-2：ページ構成自体は簡潔で見やすい → 3セクション構成をそのまま維持し他ページ改善の目標形にする → 情報量の少なさではなく「構造の分かりやすさ」を他ページ（特にブランドライブラリ）へ横展開する参照モデルとして活用",
    ]
    s = new_slide()
    title_bar(s, "④ キャンディデート向け", "対象課題：CA-1（最重要・クイックウィン）／CA-2（Good・参照モデル）", badge="candidate_pc（最重要）")
    column_label(s, LEFT_X, "AsIs（現状）", RED)
    column_label(s, RIGHT_X, "ToBe（改善後）", BLUE)

    img_top = COL_TOP + Inches(0.40)
    points_top = points_block(s, points, height=Inches(1.15))
    note_h = Inches(0.6)
    content_bottom = points_top - Inches(0.1)
    legend_h = Inches(0.68)
    left_img_h = content_bottom - img_top - legend_h - Inches(0.08)
    render_rows(s, LEFT_X, img_top, LEFT_W, left_img_h, [([(f"{ASSETS}/asis_candidate.png", "")], 1)])
    legend_block(s, LEFT_X, img_top + left_img_h + Inches(0.1), LEFT_W, legend, title="AsIs 該当箇所")

    tobe_img_h = content_bottom - img_top - note_h - Inches(0.08)
    render_rows(s, RIGHT_X, img_top, RIGHT_W, tobe_img_h, [([(f"{ASSETS}/tobe_candidate.png", "")], 1)])
    note_top = img_top + tobe_img_h + Inches(0.08)
    add_rect(s, RIGHT_X, note_top, RIGHT_W, note_h, fill=RGBColor(0xE9, 0xF3, 0xE9), line_color=RGBColor(0x3a, 0x7d, 0x3a), line_w=1.0)
    add_text(s, RIGHT_X + Inches(0.12), note_top + Inches(0.07), RIGHT_W - Inches(0.24), note_h - Inches(0.14),
              "クイックウィン：L1（既存パーツ）・費用ゼロで即着手可能。docs/02 §5・§8の最優先クイックウィンに一致。",
              size=9.3, bold=True, color=RGBColor(0x2e, 0x63, 0x2e), line_spacing=1.1)
    footer(s, "④ キャンディデート向け")


# ===========================================================================
# Slide 8: パートナー
# ===========================================================================
def slide_partner():
    legend = [
        ("①", False, "P-1　ASPチームだけ画像が読み込まれず体裁が崩れる"),
        ("②", False, "P-2　長尺スクロールなのにページ内ジャンプがない"),
    ]
    points = [
        "P-1：Akkodis ASP Teamのみ画像枠が破綻し体裁が揃わない → 現物確認の上、画像差し替えまたは枠削除で他2チームと構造を統一 → 3パートナーの見え方の一貫性を回復（低コスト・クイックウィン）",
        "P-2：長尺ページなのにページ内ナビがなくクライアントページのアンカー導線と不統一 → 3パートナーへのページ内目次を新設 → 目的パートナーへの直接到達が可能に（G-1是正の一環）",
    ]
    compare_slide(
        "⑤ グローバルブランドパートナーシップ", "対象課題：P-1／P-2", "partner_pc",
        [([(f"{ASSETS}/asis_partner_1.png", "Crop1（①）")], 0.55),
         ([(f"{ASSETS}/asis_partner_thumbnail.png", "全体縮尺（②）")], 0.45)],
        [(f"{ASSETS}/tobe_partner.png", "")],
        legend, points, points_height=Inches(1.15),
    )


# ===========================================================================
# Slide 9: 横断課題 G-1〜G-4
# ===========================================================================
def slide_cross():
    s = new_slide()
    title_bar(s, "横断課題サマリー（G-1〜G-4）", "全ページ共通テンプレート化で解決する4つの構造課題", badge="全ページ横断")

    add_text(s, Inches(0.35), Inches(1.0), Inches(6), Inches(0.24), "AsIs：各ページから代表箇所を抜粋", size=11, bold=True, color=NAVY)
    items = [
        (f"{ASSETS}/asis_g1_client.png", "G-1｜client_pc：アンカー有（Good）"),
        (f"{ASSETS}/asis_g1_brand.png", "G-1｜brand_pc：アンカー無（該当なし）"),
        (f"{ASSETS}/asis_g2_search.png", "G-2｜検索ボックス（絞り込み無）"),
        (f"{ASSETS}/asis_g3_top.png", "G-3｜top_pc：業務依頼フォーム"),
        (f"{ASSETS}/asis_g3_brand.png", "G-3｜brand_pc：Marketing部依頼フォーム"),
        (f"{ASSETS}/asis_g3_candidate.png", "G-3｜candidate_pc：各チーム連絡"),
        (f"{ASSETS}/asis_g4_nav.png", "G-4｜ナビ末尾「編集のナビゲーション」"),
    ]
    cols = 4
    cw = Inches(1.5)
    ch = Inches(0.66)
    gx, gy = Inches(0.35), Inches(1.26)
    row_pitch = ch + Inches(0.34)
    for i, (path, cap) in enumerate(items):
        cx = gx + Emu(int(cw + Inches(0.05)) * (i % cols))
        cy = gy + Emu(int(row_pitch) * (i // cols))
        add_picture_fit(s, path, cx, cy, cw, ch, valign="middle")
        add_rect(s, cx, cy, cw, ch, fill=None, line_color=RGBColor(0xC7,0xCC,0xD1), line_w=0.5)
        add_text(s, cx, cy + ch + Inches(0.01), cw, Inches(0.3), cap, size=7.3, color=MIDGRAY, line_spacing=0.95)

    ry = gy + Emu(int(row_pitch) * 2) + Inches(0.06)
    add_text(s, Inches(0.35), ry, Inches(6), Inches(0.24), "ToBe：全ページ共通テンプレート化（§6-2）", size=11, bold=True, color=NAVY)
    tobe_items = [
        "[検索ボックス web part]（G-2対応の核）：全ページヘッダー直下・共通位置に常設。Microsoft Search＋[ハイライトされたコンテンツ web part]の動的クエリを組合せ。高度ファセットはL2/L3のため将来拡張",
        "[ページ内リンク／セクションアンカー web part]（G-1対応）：全ページ冒頭にページ内目次を置くフォーマットを統一",
        "帯見出しのビジュアル統一（G-1／B-6共通）：黄色帯＋黄色マーカーの二重表現をやめ帯見出し1種類に統一",
        "「窓口案内」の共通化（G-3対応）：業務依頼フォーム／ブランドレビュー手順／各チーム連絡先を共通フォーマットのクイックリンクブロックとして各ページ末尾または目次に集約表示",
        "G-4：デザイン設計ではなくSharePoint権限・表示設定の確認事項。次ステップF0（テナント棚卸し）で確認する運用課題として扱う",
    ]
    add_multirun_bullet(s, Inches(0.35), ry + Inches(0.28), Inches(12.6), Inches(1.6), tobe_items, size=9.2, gap_pt=3)

    points = [
        "G-2（最重要）：標準検索はあるがページ内の絞り込み・ファセットがなく「見つからない→メール検索」の核心原因 → 全ページ共通位置に検索ボックス＋ハイライトされたコンテンツを設置 → 発見性の底上げ。将来的にファセット（L3）・AIアシスタント（L2）へ段階拡張",
        "G-1：ページ間で構成（アンカー有無・装飾・密度）が不統一で回遊のたび探し方を学び直す必要がある → ページ内目次・帯見出しを全ページ共通テンプレート化 → 一貫した探索体験",
        "G-3：窓口・相談導線が各所に分散 → 共通フォーマットの窓口案内ブロックを各ページに配置 → どの窓口を使うべきかが一目で分かる",
        "G-4：ナビ末尾に「編集 のナビゲーション」が露出（編集者ビュー由来と見られ要確認） → デザイン変更ではなくSharePoint表示設定の確認事項としてF0（次ステップ）に回す",
    ]
    points_block(s, points, height=Inches(1.75))
    footer(s, "横断課題サマリー")


# ===========================================================================
# Slide 10: 次ステップ
# ===========================================================================
def slide_next():
    s = new_slide()
    title_bar(s, "次ステップ", "フェーズ計画（F0〜F4）と社長判断3点（docs/02 §6）", badge="次アクション")

    phases = [
        ("F0", "テナント棚卸し", "M365 Copilot／Copilot Studio／SharePoint Premium／Power Platformのライセンス確認、SPFxカスタムコード許可可否、IT窓口特定", "ops-manager（社長経由でIT窓口）"),
        ("F1", "要件定義", "課題→機能要件化、対象利用者・権限区分、AIスコープ確定", "ops-manager／marketer"),
        ("F2", "SharePoint情報設計", "サイト構成・ナビ/IA・ページ設計・メタデータ列/コンテンツタイプ・権限方針・検索スキーマ方針", "engineer／designer"),
        ("F3", "L1 PoC／MVP構築", "クイックウィン実装、ブランドライブラリ再構成、Microsoft Search+web part配置、Lists/Forms受付、メタデータ適用", "engineer"),
        ("F4", "L2拡張（条件付き）", "F0結果に応じAI-6メタデータ自動付与／AI-2 Copilotアシスタント／AI-5申請自動化を段階投入", "engineer／qa-reviewer"),
    ]
    x = Inches(0.35)
    tw = Inches(2.44)
    y = Inches(1.15)
    for i, (code, name, desc, owner) in enumerate(phases):
        cx = x + Emu(int(tw) * i)
        fill = YELLOW if code == "F0" else LIGHTBG
        add_rect(s, cx, y, tw - Inches(0.06), Inches(0.42), fill=NAVY)
        add_text(s, cx, y, tw - Inches(0.06), Inches(0.42), f"{code}　{name}", size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, cx, y + Inches(0.42), tw - Inches(0.06), Inches(1.55), fill=RGBColor(0xFF,0xF7,0xE2) if code=="F0" else LIGHTBG, line_color=RGBColor(0xC7,0xCC,0xD1), line_w=0.75)
        add_text(s, cx + Inches(0.08), y + Inches(0.5), tw - Inches(0.2), Inches(1.15), desc, size=8.2, color=DARKGRAY, line_spacing=1.05)
        add_text(s, cx + Inches(0.08), y + Inches(1.72), tw - Inches(0.2), Inches(0.22), f"主担当: {owner}", size=7.3, color=MIDGRAY)
    for i in range(len(phases) - 1):
        cx = x + Emu(int(tw) * i)
        add_text(s, cx + tw - Inches(0.13), y + Inches(0.06), Inches(0.13), Inches(0.3), "→", size=13, bold=True,
                  color=YELLOW, align=PP_ALIGN.CENTER)

    add_text(s, Inches(0.35), Inches(3.1), Inches(6), Inches(0.24), "F0はF1より先（何が作れるか未確定のまま設計に入らない）", size=9.5, italic=True, color=MIDGRAY)

    add_text(s, Inches(0.35), Inches(3.55), Inches(6), Inches(0.26), "社長にご判断・ご確認いただきたい点（3点）", size=13, bold=True, color=NAVY)
    decisions = [
        ("【要判断】", "MVPをL1の4点に絞って先行してよいか", "直リンク化(CA-1)＋ブランドライブラリ再構成(B-1/B-2)＋Microsoft Search(G-2)＋メタデータ設計の4点を先行する案。推奨：絞る"),
        ("【要確認】", "テナントのライセンス実態の確認（IT窓口の特定）", "M365 Copilot／Copilot Studio／SharePoint Premium／Power Platformプレミアムの保有・付与範囲。未確定だとAI6機能の可否が決まらない（最大のブロッカー）"),
        ("【要判断】", "クイックウィン（直リンク化・費用ゼロ・L1）をF0完了を待たず先行着手してよいか", ""),
    ]
    dy = Inches(3.85)
    for tag, name, desc in decisions:
        add_rect(s, Inches(0.35), dy, Inches(1.0), Inches(0.9 if desc else 0.4), fill=NAVY)
        add_text(s, Inches(0.35), dy, Inches(1.0), Inches(0.9 if desc else 0.4), tag, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(1.45), dy, Inches(11.3), Inches(0.24), name, size=10.5, bold=True, color=DARKGRAY)
        if desc:
            add_text(s, Inches(1.45), dy + Inches(0.25), Inches(11.3), Inches(0.6), desc, size=8.8, color=MIDGRAY, line_spacing=1.1)
        dy += Inches(0.9 if desc else 0.4) + Inches(0.12)

    footer(s, "次ステップ")


if __name__ == "__main__":
    slide_title()
    slide_exec_summary()
    slide_top()
    slide_brand_1()
    slide_brand_2()
    slide_client()
    slide_candidate()
    slide_partner()
    slide_cross()
    slide_next()
    prs.save(OUT_PATH)
    print("saved", OUT_PATH, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))

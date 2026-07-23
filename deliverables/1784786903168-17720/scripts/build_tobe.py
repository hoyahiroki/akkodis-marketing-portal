#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ToBe ワイヤーフレーム生成スクリプト
- HTML/CSSで構造図(灰色矩形+ラベル)を組み、Playwright(Chromium)でPNG化
- 配色・装飾は作り込まない(designer-uiが後工程で配色を当てる前提)
- L1範囲外要素は点線枠+「将来拡張(L2以上)」ラベルで区別
"""
import os
from playwright.sync_api import sync_playwright

ROOT = "/workspace/akkodis-marketing-portal"
OUT = os.path.join(ROOT, "deliverables/1784786903168-17720/assets")
os.makedirs(OUT, exist_ok=True)

CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

CSS = """
* { box-sizing: border-box; margin:0; padding:0; }
body { font-family: "IPAGothic","IPAPGothic",sans-serif; background:#eef0f2; padding:18px; width:__WIDTH__px; }
.page { background:#fff; border:2px solid #333; }
.blk { border:2px solid #7a7a7a; background:#f7f7f8; margin:10px; padding:10px 12px; }
.blk-label { font-size:13px; font-weight:bold; color:#222; margin-bottom:4px; }
.blk-part { display:inline-block; font-size:11px; color:#444; background:#e2e4e7; border:1px solid #9a9a9a; padding:1px 6px; border-radius:3px; margin-bottom:4px; font-family: monospace; }
.blk-note { font-size:11px; color:#555; margin-top:3px; }
.row { display:flex; gap:10px; }
.col { flex:1; }
.header { background:#dfe2e6; border:2px solid #555; margin:10px; padding:8px 12px; display:flex; align-items:center; justify-content:space-between; }
.header .searchbox { flex:1; max-width:340px; background:#fff; border:1px solid #888; border-radius:3px; padding:5px 8px; font-size:12px; color:#666; margin-left:14px; }
.header .brand { font-size:13px; font-weight:bold; color:#222; }
.hero { margin:10px; height:70px; background:repeating-linear-gradient(45deg,#d7d9db,#d7d9db 10px,#c9cbce 10px,#c9cbce 20px); border:2px solid #7a7a7a; display:flex; align-items:center; padding:0 16px; font-weight:bold; font-size:15px; color:#333; }
.toc { margin:10px; border:2px solid #555; background:#fff; padding:8px 12px; }
.toc .blk-part { margin-bottom:6px; }
.toc-list { display:flex; flex-wrap:wrap; gap:8px; }
.toc-item { background:#eef2fb; border:1px solid #5a7fc7; color:#2c4a91; font-size:11.5px; padding:4px 9px; border-radius:14px; }
.toc-item:before { content:"↳ "; }
.quicklinks { display:flex; gap:10px; }
.qcard { flex:1; border:2px solid #7a7a7a; background:#fff; padding:10px; text-align:center; }
.qcard .icon { width:28px; height:28px; background:#c9cbce; margin:0 auto 6px; border-radius:4px; }
.qcard .t { font-size:12px; font-weight:bold; color:#222; }
.qcard .d { font-size:10.5px; color:#666; margin-top:4px; }
.library { border:2px solid #555; background:#fff; padding:10px 12px; }
.lib-meta { display:flex; gap:6px; margin-bottom:8px; flex-wrap:wrap; }
.lib-meta span { font-size:10px; background:#e2e4e7; border:1px solid #9a9a9a; padding:2px 6px; border-radius:3px; color:#333; }
.lib-cats { display:flex; gap:8px; flex-wrap:wrap; }
.lib-cat { flex:1; min-width:150px; border:1.5px dashed #888; padding:8px; }
.lib-cat .n { font-size:11.5px; font-weight:bold; margin-bottom:4px; color:#222; }
.lib-cat ul { list-style:none; }
.lib-cat li { font-size:10px; color:#555; padding:2px 0 2px 10px; position:relative; }
.lib-cat li:before { content:"▪"; position:absolute; left:0; color:#999; font-size:8px; }
.steps { border:2px solid #7a7a7a; background:#fff; padding:10px 12px; }
.steps ol { margin-left:18px; font-size:11.5px; color:#333; line-height:1.9; }
.dashed { border:2px dashed #b5762a; background:#fbf3e8; margin:10px; padding:10px 12px; }
.dashed .blk-label { color:#8a5518; }
.future-tag { display:inline-block; font-size:10.5px; color:#8a5518; background:#f3dfc0; border:1px solid #b5762a; padding:1px 7px; border-radius:3px; margin-bottom:4px; }
.textblk { margin:10px; padding:8px 12px; border:1.5px solid #aab; background:#fbfbfd; font-size:11.5px; color:#333; }
.footer-tag { text-align:center; font-size:11px; color:#888; padding:8px; }
h1.title { font-size:16px; padding:10px 14px 0; color:#111; }
"""

HTML_TMPL = """<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head>
<body><div class="page">{body}</div></body></html>"""


def block(label, part=None, note=None, extra_class=""):
    part_html = f'<div class="blk-part">{part}</div><br/>' if part else ""
    note_html = f'<div class="blk-note">{note}</div>' if note else ""
    return f'<div class="blk {extra_class}"><div class="blk-label">{part_html}{label}</div>{note_html}</div>'


def header(site="マーケティングポータル", extra=""):
    return f'''<div class="header">
      <div class="brand">☰ {site}</div>
      <div class="searchbox">🔍 [検索ボックス web part] このサイトを検索{(" ・ "+extra) if extra else ""}</div>
    </div>'''


def hero(title):
    return f'<div class="hero">{title}（ヒーロー・現状維持）</div>'


def toc(items, part="[ページ内リンク／セクションアンカー web part]"):
    its = "".join(f'<div class="toc-item">{i}</div>' for i in items)
    return f'''<div class="toc"><div class="blk-part">{part}</div><div class="toc-list">{its}</div></div>'''


def quicklinks(cards, part="[クイックリンク web part]", title=None):
    cs = "".join(f'<div class="qcard"><div class="icon"></div><div class="t">{c[0]}</div><div class="d">{c[1]}</div></div>' for c in cards)
    t = f'<div class="blk-label">{title}</div>' if title else ""
    return f'<div class="blk"><div class="blk-part">{part}</div>{t}<div class="quicklinks">{cs}</div></div>'


def library(title, part, meta, cats, note=None):
    metas = "".join(f"<span>{m}</span>" for m in meta)
    cathtml = "".join(
        f'<div class="lib-cat"><div class="n">{c[0]}</div><ul>{"".join(f"<li>{x}</li>" for x in c[1])}</ul></div>'
        for c in cats
    )
    note_html = f'<div class="blk-note">{note}</div>' if note else ""
    return f'''<div class="blk"><div class="blk-part">{part}</div><div class="blk-label">{title}</div>
    <div class="library"><div class="lib-meta">{metas}</div><div class="lib-cats">{cathtml}</div></div>{note_html}</div>'''


def steps(title, items, part="[テキストWebパーツ／番号付きリスト]", note=None):
    lis = "".join(f"<li>{i}</li>" for i in items)
    note_html = f'<div class="blk-note">{note}</div>' if note else ""
    return f'<div class="blk"><div class="blk-part">{part}</div><div class="blk-label">{title}</div><div class="steps"><ol>{lis}</ol></div>{note_html}</div>'


def dashed(label, note=None):
    note_html = f'<div class="blk-note">{note}</div>' if note else ""
    return f'<div class="dashed"><span class="future-tag">将来拡張（L2以上）</span><div class="blk-label">{label}</div>{note_html}</div>'


def textblk(text):
    return f'<div class="textblk">{text}</div>'


def footer():
    return '<div class="footer-tag">— ページ末尾: 共通「窓口案内」クイックリンクブロック（G-3対応・全ページ共通フォーマット） —</div>'


def render(name, blocks, width=760):
    css = CSS.replace("__WIDTH__", str(width))
    body = "".join(blocks)
    html = HTML_TMPL.format(css=css, body=body)
    html_path = os.path.join(OUT, f"tmp_{name}.html")
    with open(html_path, "w") as f:
        f.write(html)
    png_path = os.path.join(OUT, f"tobe_{name}.png")
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME, args=["--no-sandbox"])
        page = b.new_page(viewport={"width": width + 40, "height": 800})
        page.goto("file://" + html_path)
        page.wait_for_timeout(80)
        el = page.query_selector(".page")
        el.screenshot(path=png_path)
        b.close()
    os.remove(html_path)
    print("saved", png_path)


# ---------------------------------------------------------------------------
# 1. トップ
# ---------------------------------------------------------------------------
def build_top():
    blocks = [
        header(),
        hero("マーケティングポータル"),
        quicklinks(
            [("キャンディデート向け", "説明文1行（今後定義）"), ("クライアント向け", "説明文1行（今後定義）"),
             ("ブランドライブラリ", "説明文1行（今後定義）"), ("グローバルブランドパートナーシップ", "説明文1行（今後定義）")],
            title="4カテゴリカード（説明文欄を追加＝T-3対応）"
        ),
        quicklinks(
            [("業務依頼フォーム", "（使い分けは今後定義）"), ("ブランドレビューの手順", "（使い分けは今後定義）")],
            title="「よく使う窓口」セクション（新設・T-4対応）"
        ) + textblk("テキストWebパーツ: 2つの使い分けを1行で説明（文言は今後定義）"),
        block("「Webサイト流入分析」セクション：PC=現状の埋め込みダッシュボード維持", part="[埋め込みWebパーツ web part]",
              note="SP表示のみ→ 主要指標サマリー＋「詳細はPCでご確認ください」に切替（T-1対応）", extra_class=""),
        block("「組織体制」セクション（配置順は維持）", part="[組織図 web part]",
              note="目次アンカー「組織体制を見る」から直接ジャンプ可能に（T-2対応）"),
        footer(),
    ]
    render("top", blocks)


# ---------------------------------------------------------------------------
# 2. ブランドライブラリ (2スライド分の構成要素を1枚のワイヤーとして生成し、
#    pptx側で上下2分割して2スライドに割り付ける)
# ---------------------------------------------------------------------------
def build_brand_part1():
    blocks = [
        header(),
        hero("ブランドライブラリ"),
        toc(["ブランドの使用について", "ガイドライン一覧", "ロゴ掲載申請", "フォント", "画像素材", "各種テンプレート", "SNSガイドライン"]),
        block("「ブランドの使用について」：黄色マーカー3種→帯見出し1種類に統一（B-6対応）", part="[テキストWebパーツ]",
              note="「ブランドレビュー依頼の前にご確認ください」「注意事項（社外秘）」は折りたたみ可能ブロックに"),
        quicklinks(
            [("①", "ラベル（今後定義）"), ("②", "ラベル（今後定義）"), ("③", "ラベル（今後定義）")],
            title="アイコンカード3枚にラベル追加（B-5対応）"
        ),
        library(
            "「ガイドラインライブラリ」セクション（B-1／B-2の核心・最重要）",
            "[ドキュメントライブラリ web part] + [ハイライトされたコンテンツ web part]",
            ["メタ列: 種別", "言語", "版", "形式"],
            [
                ("全社ブランド規定", ["Akkodis_ブランドガイドライン_V3", "AKKODiS新ビジョンガイドライン"]),
                ("制作物別ガイドライン", ["ビデオ", "フォトグラフィー", "AI生成画像の使用"]),
                ("表記・用語", ["文字表記", "グローバル用語集", "ブランドボイス"]),
                ("Akkodis Intelligence", ["ガイドライン", "Q&A"]),
            ],
            note="既存10本を目的軸4カテゴリに束ね直し、ビュー切替でコンテンツタイプ別フィルタも可能に（B-2対応）"
        ),
    ]
    render("brand_1", blocks)


def build_brand_part2():
    blocks = [
        header(),
        steps(
            "「ロゴ掲載申請」セクション",
            ["Marketing部依頼フォーム", "Teams連絡",
             "Salesforceのサービスリクエストから押印申請", "回収",
             "送付", "HP掲載"],
            part="[ドキュメントライブラリ web part]（同意書テンプレへ直リンク）＋番号付きリスト",
            note="B-3対応：番号付き手順で可視化・関連テンプレ直リンク。3システム構成自体は本MVP範囲外（自動化はL2）"
        ),
        block("「フォント」セクション（現状のテキスト構成を維持、目次から到達可能に）", part="[テキストWebパーツ]"),
        block("「画像素材」セクション（現状のテキスト構成を維持、目次から到達可能に）", part="[テキストWebパーツ]"),
        library(
            "「各種テンプレート」帯",
            "[ハイライトされたコンテンツ web part]",
            ["コンテンツタイプで絞り込み"],
            [("テンプレート", ["PowerPoint", "Word", "メール署名", "Forms", "バーチャル背景"])],
        ),
        quicklinks([("ソーシャルメディアガイドライン", "現状維持")], title="「ソーシャルメディアガイドライン」帯"),
        dashed("AI検索の高度ファセット／Copilot要約等", note="L1範囲外のためMVPワイヤー本体には含めない"),
        footer(),
    ]
    render("brand_2", blocks)


# ---------------------------------------------------------------------------
# 3. クライアント向け
# ---------------------------------------------------------------------------
def build_client():
    blocks = [
        header(),
        hero("クライアント向け"),
        toc(["AKKODiS企業向けWebページ", "公開資料ライブラリ", "メルマガ配信スケジュール", "資料を探す（新規）"]),
        block("このページ内資料検索（新設）", part="[検索ボックス web part]"),
        quicklinks([("セミナー", ""), ("活用事例", ""), ("Insights", ""), ("サービス", "")],
                   title="「AKKODiS企業向けWebページ」セクション（現状維持）"),
        library(
            "「公開資料ライブラリ」セクション（旧・ホームページ公開資料／C-1・C-2対応の核）",
            "[ドキュメントライブラリ web part] + [ハイライトされたコンテンツ web part]",
            ["カテゴリ(共通/ソリューション/アカデミー/事例/お役立ち)", "言語", "更新日"],
            [
                ("共通", ["会社案内(日/英)", "（他資料は今後定義）"]),
                ("ソリューション", ["紹介資料 約11本"]),
                ("アカデミー", ["紹介資料 約8本"]),
                ("お役立ち資料", ["30本超（日本語版/英語版は言語メタデータで1エントリに集約）"]),
            ],
            note="日本語版／英語版が対のものは「言語」メタデータで統合・切替表示（C-2対応）"
        ),
        quicklinks([("メルマガ配信スケジュール", "※社外秘・現状維持")], title="「メルマガ配信スケジュール」セクション"),
        footer(),
    ]
    render("client", blocks, width=760)


# ---------------------------------------------------------------------------
# 4. キャンディデート向け
# ---------------------------------------------------------------------------
def build_candidate():
    blocks = [
        header(),
        hero("キャンディデート向けページ"),
        quicklinks([("AKKODiS People", ""), ("リクルーティングムービー", "")], title="「AKKODiS公式HP」セクション（現状維持）"),
        library(
            "「マスター資料」セクション（CA-1対応の核・最重要）",
            "[ドキュメントライブラリ web part] + [クイックリンク web part]（直リンク）",
            ["資料オーナー／問い合わせ先", "更新日"],
            [("会社説明資料（直リンク）", ["中途採用向け", "新卒採用向け", "障がい者採用向け"])],
            note="「各チームまで直接ご連絡ください」の行き止まりを廃止。担当チーム名は資料メタデータ「問い合わせ先」として保持。"
                 "[ハイライトされたコンテンツ web part]で最新版を自動前面表示"
        ),
        quicklinks([("X", ""), ("Facebook", ""), ("LinkedIn", ""), ("Youtube", "")], title="「Social Media」セクション（現状維持）"),
        footer(),
    ]
    render("candidate", blocks, width=700)


# ---------------------------------------------------------------------------
# 5. グローバルブランドパートナーシップ
# ---------------------------------------------------------------------------
def build_partner():
    blocks = [
        header(),
        hero("グローバルブランドパートナーシップ（リード文含む）"),
        quicklinks([("Mercedes-AMG PETRONAS", ""), ("Akkodis ASP Team", ""), ("Stade Toulousain", "")],
                   title="3パートナーのサムネイルカード（現状維持・リンク先強化）"),
        toc(["Mercedes-AMG PETRONAS Formula One Team", "Akkodis ASP Team", "Stade Toulousain"], part="[ページ内リンク／セクションアンカー web part]（新設・P-2対応）"),
        block("各パートナーセクション（概要／共通価値／パートナーシップビデオ・現状維持）", part="[テキストWebパーツ]+[埋め込み動画]", extra_class=""),
        block("Akkodis ASP Teamセクションの画像枠是正（P-1対応）", part="[画像 web part]",
              note="現物確認の上、正しい画像へ差し替え、または他2チームと同じ構造に揃えて枠を削除（差し替え画像の内容は指定しない）"),
        footer(),
    ]
    render("partner", blocks, width=740)


if __name__ == "__main__":
    build_top()
    build_brand_part1()
    build_brand_part2()
    build_client()
    build_candidate()
    build_partner()
    print("all done")

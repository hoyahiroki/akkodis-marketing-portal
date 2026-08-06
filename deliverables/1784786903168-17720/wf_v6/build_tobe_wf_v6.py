#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ToBe ワイヤーフレーム（清書版）生成スクリプト

build_tobe_wf_v5.py を土台に、「ワイヤーフレームは構成だけを示す」方針で作り直したもの。
- 画像に描くのは: ページ名 / セクション番号バッジ / セクション見出し / 使用 Web パーツ名 /
  各セクションの中身の模式表現（カード・リンク箇条書き・3ステップ導線など）のみ。
- 説明文・検討経緯・注記の類は画像から全て外し、セクション番号に紐づけて comments.json に出力する。
- 設計内容そのものは変更していない。変えたのは見せ方（レイアウト）と文章（清書）だけ。

出力:
  {top,brand,client,candidate,partner}.png … 5枚とも同一ピクセル幅
  comments.json … セクション番号ごとの説明文（画像の外に置くコメント表のデータ源）
"""
import json
import os

from playwright.sync_api import sync_playwright

OUT = "/workspace/akkodis-marketing-portal/deliverables/1784786903168-17720/wf_v6"
os.makedirs(OUT, exist_ok=True)

CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

WIDTH = 1160          # 本文幅（全ページ共通・レンダリング後は device_scale_factor 倍）
SCALE = 2
TOTAL_PAGES = 5

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
CSS = """
* { box-sizing: border-box; margin:0; padding:0; }
body {
  font-family: "IPAGothic","IPAPGothic","Noto Sans CJK JP",sans-serif;
  background:#e7e9ec; padding:0 0 34px 0; width: __WIDTH__px;
}
.wrap { width:100%; }

/* ページ名見出し */
.titlebar {
  background:#1c2733; color:#fff; padding:18px 26px;
  display:flex; justify-content:space-between; align-items:baseline;
}
.titlebar .name { font-size:24px; font-weight:bold; letter-spacing:.5px; }
.titlebar .sub { font-size:13px; color:#b8c2cf; }

.page { background:#fff; border:2px solid #333; padding:6px 0 10px 0; }

/* 番号バッジ付きブロックの共通レイアウト */
.blk { display:flex; align-items:flex-start; gap:14px; margin:16px 20px; }
.blk .numbadge {
  flex:0 0 46px; width:46px; height:46px; border-radius:50%;
  background:#1c2733; color:#fff; font-size:20px; font-weight:bold;
  font-family:"DejaVu Sans",Arial,Helvetica,sans-serif; letter-spacing:0;
  display:flex; align-items:center; justify-content:center;
  border:2px solid #1c2733;
}
.blk.header .numbadge { background:#fff; color:#1c2733; }
.blk .body { flex:1; min-width:0; }

/* 共通ヘッダー（サイト内検索／サイト共通エージェント） */
.headerband { border:2px solid #55606b; background:#f2f4f6; }
.headerband .row {
  display:flex; align-items:center; gap:12px; padding:11px 18px; font-size:14px; color:#2a2f36;
}
.headerband .row + .row { border-top:1.5px dashed #9aa2ab; }
.headerband .tag {
  background:#55606b; color:#fff; padding:3px 10px; border-radius:3px;
  font-size:11.5px; font-weight:bold; white-space:nowrap;
}
.headerband .tag.ai { background:#2f6b3f; }

/* セクション枠 */
.sec { border:2px solid #7a7a7a; background:#f7f7f8; padding:15px 20px; }
.sec.tone-a { background:#fff4e0; border-color:#c58a2e; }
.sec.tone-b { background:#fdeceb; border-color:#c0392b; }
.sec.dashed { border-style:dashed; background:#fbf3e8; border-color:#b5762a; }
.sec .sechead { display:flex; justify-content:space-between; align-items:flex-start; gap:14px; }
.sec .title { font-size:19px; font-weight:bold; color:#181b1f; margin:7px 0 10px; }
.parts { display:flex; flex-wrap:wrap; gap:6px; }
.part-tag {
  display:inline-block; font-size:12.5px;
  background:#e2e4e8; border:1px solid #9a9fa6; color:#33383e;
  padding:3px 10px; border-radius:4px;
}
.issue {
  font-size:12px; font-weight:bold; padding:3px 9px; border-radius:3px;
  border:1px solid #5c85b8; background:#e6f0fb; color:#204a78; white-space:nowrap;
}

/* ヒーロー */
.hero {
  min-height:74px; border:2px solid #7a7a7a;
  background:repeating-linear-gradient(45deg,#dfe1e3,#dfe1e3 12px,#d0d2d5 12px,#d0d2d5 24px);
  display:flex; flex-direction:column; justify-content:center; padding:16px 22px;
}
.hero .t { font-size:20px; font-weight:bold; color:#20242a; }
.hero .s { font-size:13.5px; color:#4a4f57; margin-top:5px; }

/* 目次 */
.toc-items { display:flex; flex-wrap:wrap; gap:10px; }
.toc-item {
  background:#fff; border:1.5px solid #3b5bb0; border-radius:20px;
  padding:9px 16px; font-size:13.5px; color:#1c3570; line-height:1.5;
}
.toc-item b { color:#0d47a1; }

/* カード群 */
.cards { display:grid; grid-template-columns: repeat(3, 1fr); gap:13px; }
.cards.cols4 { grid-template-columns: repeat(4, 1fr); }
.card { border:2px solid #8a8f96; background:#fff; padding:13px 13px; }
.card .icon { width:32px; height:32px; background:#c9cbce; border-radius:6px; margin-bottom:8px; }
.card .t { font-size:14.5px; font-weight:bold; color:#1e2126; margin-bottom:6px; }
.card .d { font-size:12.5px; color:#555; line-height:1.6; }

/* 縦積み3ステップ導線 */
.flowV { display:flex; flex-direction:column; }
.flowV .stepRow { display:flex; gap:14px; align-items:stretch; }
.flowV .stepnum {
  flex:0 0 30px; width:30px; height:30px; border-radius:50%;
  background:#555; color:#fff; font-size:14px; font-weight:bold;
  display:flex; align-items:center; justify-content:center;
}
.flowV .stepbox { flex:1; min-width:0; }
.flowV .arrowDown { margin-left:13px; padding:3px 0; font-size:20px; color:#8a8f96; line-height:1; }
.flowV .guidebox { border:2px solid #8a8f96; background:#fff; padding:14px 18px; }
.flowV .guidebox .t { font-size:14.5px; font-weight:bold; color:#1e2126; }
.flowV .formbox {
  border:2px solid #2b3a55; background:#eef1f6; padding:14px 18px;
  display:flex; align-items:center;
}
.cta-primary {
  display:inline-block; background:#2b3a55; color:#fff; font-size:15.5px; font-weight:bold;
  padding:12px 24px; border-radius:5px; border:2px solid #1c2733; line-height:1.5;
}

/* AI-1 埋め込みボックス */
.ai1box { border:2px solid #1c5f8a; background:#eaf3fa; padding:13px 18px; }
.ai1box .tag {
  display:inline-block; font-size:11.5px; font-weight:bold; color:#fff; background:#1c5f8a;
  padding:2px 9px; border-radius:4px; margin-bottom:7px;
}
.ai1box .t { font-size:15px; font-weight:bold; color:#123b54; }

/* 2カラム */
.cols2 { display:flex; gap:16px; }
.col { flex:1; min-width:0; border:1.5px solid #9a9fa6; background:#fff; padding:13px 16px; }
.col .t { font-size:14.5px; font-weight:bold; margin-bottom:8px; color:#1e2126; }
.col ol { margin-left:20px; font-size:13px; line-height:1.95; color:#2b2b2b; }
.col .link { margin-top:10px; font-size:12.5px; color:#2b3a55; border-top:1px dashed #aaa; padding-top:8px; }

/* リンク一覧（テキスト Web パーツの見出し＋リンク箇条書き） */
.linkgroups { display:grid; grid-template-columns: repeat(2, 1fr); gap:13px; }
.linkgroups.g1 { grid-template-columns: 1fr; }
.linkgroup { border:1.5px solid #9a9fa6; background:#fbfbfc; padding:12px 16px; }
.linkgroup .n { font-size:13.5px; font-weight:bold; margin-bottom:7px; color:#1e2126; }
.linkgroup ul { list-style:none; }
.linkgroup li {
  font-size:12.5px; color:#1c5f8a; padding:4px 0 4px 15px; position:relative;
  line-height:1.65; text-decoration:underline;
}
.linkgroup li:before { content:"—"; position:absolute; left:0; color:#8a8f96; text-decoration:none; }

/* 折りたたみセクション */
.collapsible { border:2px solid #55606b; background:#fff; }
.collapse-head {
  background:#e8eaed; color:#2a2f36; font-size:13px; font-weight:bold;
  padding:9px 16px; border-bottom:2px solid #55606b;
}
.collapse-body { padding:13px 16px; }

/* サブブロック（現行文言の掲出） */
.subblock { border-left:6px solid #999; padding:11px 16px; background:#fafafa; }
.subblock .st { font-size:13.5px; font-weight:bold; margin-bottom:6px; color:#222; }
.subblock .sd { font-size:13px; color:#3a3a3a; line-height:1.75; }

/* メタ列 */
.lib-meta { display:flex; gap:8px; flex-wrap:wrap; }
.lib-meta span {
  font-size:12px; background:#e2e4e8; border:1px solid #9a9fa6; color:#333;
  padding:4px 10px; border-radius:3px;
}

/* リスト Web パーツ（副案の模式図） */
.listpart { border:1.5px solid #8a8f96; background:#fff; }
.listpart .filterbar {
  display:flex; gap:8px; flex-wrap:wrap; padding:9px 13px; background:#eef1f6;
  border-bottom:1.5px solid #8a8f96; align-items:center;
}
.listpart .filterbar .lbl { font-size:11.5px; color:#444; font-weight:bold; margin-right:4px; }
.listpart .chip {
  font-size:11.5px; background:#fff; border:1.5px solid #3b5bb0; color:#1c3570;
  padding:3px 11px; border-radius:14px;
}
.listpart table { width:100%; border-collapse:collapse; }
.listpart th, .listpart td {
  border:1px solid #d3d5d8; padding:7px 11px; font-size:12.5px;
  text-align:left; vertical-align:top; line-height:1.55;
}
.listpart th { background:#e9ebee; font-size:12px; }
.listpart td.linkcol { color:#1c5f8a; text-decoration:underline; }

/* 副案（点線枠）の見出し */
.altlabel {
  display:inline-block; font-size:12px; font-weight:bold; color:#8a5518; background:#f3dfc0;
  border:1px solid #b5762a; padding:3px 10px; border-radius:4px; margin-bottom:8px;
}
.altbox { border:2.5px dashed #b5762a; background:#fdf7ef; padding:13px 16px; }
.altbox .t { font-size:14.5px; font-weight:bold; color:#6f4413; margin-bottom:9px; }

.stack > * + * { margin-top:11px; }
.tbd {
  color:#8a4b00; background:#fff3de; border:1px solid #d69a3f;
  padding:1px 7px; border-radius:3px; font-size:12px; font-weight:bold;
  text-decoration:none; display:inline-block;
}
"""

HTML_TMPL = """<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head>
<body><div class="wrap">{body}</div></body></html>"""


# ---------------------------------------------------------------------------
# 共通ビルダー
# ---------------------------------------------------------------------------
def titlebar(name, idx, total):
    return f'''<div class="titlebar">
      <div class="name">ToBe ワイヤーフレーム　｜　{name}　（{idx}/{total}）</div>
      <div class="sub">AKKODiSマーケティングポータル</div>
    </div>'''


def header_band():
    """全ページ共通ヘッダー（セクション番号 0）。ラベルのみで注記帯は置かない。"""
    return '''<div class="blk header" data-sec="0">
      <div class="numbadge">0</div>
      <div class="body headerband">
        <div class="row">
          <span class="tag">システム標準</span>
          <span>サイトヘッダー：サイト内検索（Microsoft Search）／位置固定・全ページ共通</span>
        </div>
        <div class="row">
          <span class="tag ai">実装要素</span>
          <span>サイトヘッダー：AI-2（SharePoint agent）／サイト共通エージェント・全ページ共通</span>
        </div>
      </div>
    </div>'''


def block(num, inner, extra_cls=""):
    return f'''<div class="blk" data-sec="{num}">
      <div class="numbadge">{num}</div>
      <div class="body {extra_cls}">{inner}</div>
    </div>'''


def sec(num, title, parts, inner="", tone="", issue=""):
    """番号バッジ付きセクション。parts=Webパーツ名のリスト、inner=中身の模式表現。"""
    parts_html = "".join(f'<span class="part-tag">[{p}]</span>' for p in parts)
    issue_html = f'<span class="issue">{issue}</span>' if issue else ""
    klass = f"sec {tone}".strip()
    body = f'''<div class="{klass}">
      <div class="sechead">
        <div class="parts">{parts_html}</div>
        {issue_html}
      </div>
      <div class="title">{title}</div>
      {inner}
    </div>'''
    return f'''<div class="blk" data-sec="{num}">
      <div class="numbadge">{num}</div>
      <div class="body">{body}</div>
    </div>'''


def hero(title, sub=""):
    sub_html = f'<div class="s">{sub}</div>' if sub else ""
    return f'<div class="hero"><div class="t">{title}</div>{sub_html}</div>'


def toc_items(items):
    its = "".join(
        f'<div class="toc-item">→ <b>{it["label"]}</b>　｜遷移先: {it["target"]}</div>'
        for it in items
    )
    return f'<div class="toc-items">{its}</div>'


def cards(items, cols=3):
    cls = "cards cols4" if cols == 4 else "cards"
    cs = "".join(
        f'<div class="card"><div class="icon"></div><div class="t">{t}</div><div class="d">{d}</div></div>'
        for t, d in items
    )
    return f'<div class="{cls}">{cs}</div>'


def linkgroups(groups, single_col=False):
    cls = "linkgroups g1" if single_col else "linkgroups"
    gs = "".join(
        f'<div class="linkgroup"><div class="n">{g[0]}</div>'
        f'<ul>{"".join(f"<li>{x}</li>" for x in g[1])}</ul></div>'
        for g in groups
    )
    return f'<div class="{cls}">{gs}</div>'


def collapsible(inner_html, head="▼ このセクションを折りたたみ可能にする（既定：折りたたみ）"):
    return f'''<div class="collapsible">
      <div class="collapse-head">{head}</div>
      <div class="collapse-body">{inner_html}</div>
    </div>'''


def cols2(col1, col2):
    """col = (見出し, [手順...], リンク行 or None)"""
    def render(c):
        lis = "".join(f"<li>{s}</li>" for s in c[1])
        link = f'<div class="link">— [ドキュメント ライブラリ Web パーツ] {c[2]}</div>' if c[2] else ""
        return f'<div class="col"><div class="t">{c[0]}</div><ol>{lis}</ol>{link}</div>'
    return f'<div class="cols2">{render(col1)}{render(col2)}</div>'


def ai1box():
    return '''<div class="ai1box">
      <span class="tag">実装要素</span>
      <div class="t">[埋め込み Web パーツ] AI-1（ブランドセルフチェッカー）</div>
    </div>'''


# 測定結果（ページ名 -> {セクション番号: {"y":…, "h":…}}）。render() が書き込む。
GEOM = {}


def render(name, idx, page_label, blocks):
    css = CSS.replace("__WIDTH__", str(WIDTH))
    body = titlebar(page_label, idx, TOTAL_PAGES) + '<div class="page">' + "".join(blocks) + "</div>"
    html = HTML_TMPL.format(css=css, body=body)
    html_path = os.path.join(OUT, f"_tmp_{name}.html")
    with open(html_path, "w") as f:
        f.write(html)
    png_path = os.path.join(OUT, f"{name}.png")
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME, args=["--no-sandbox"])
        page = b.new_page(viewport={"width": WIDTH, "height": 900}, device_scale_factor=SCALE)
        page.goto("file://" + html_path)
        page.wait_for_timeout(120)
        page.screenshot(path=png_path, full_page=True)
        # スクリーンショットと同一のレンダリング条件のまま、各セクションブロックの
        # 位置・高さを取得する。値は出力PNGの実ピクセル基準（device_scale_factor 適用後）。
        geom = page.evaluate(
            """(scale) => {
              const out = {};
              document.querySelectorAll('[data-sec]').forEach(el => {
                const r = el.getBoundingClientRect();
                out[el.getAttribute('data-sec')] = {
                  y: Math.round((r.top + window.scrollY) * scale),
                  h: Math.round(r.height * scale)
                };
              });
              out['__img__'] = {
                w: Math.round(document.documentElement.scrollWidth * scale),
                h: Math.round(document.documentElement.scrollHeight * scale)
              };
              return out;
            }""",
            SCALE,
        )
        b.close()
    os.remove(html_path)
    GEOM[name] = geom
    print("saved", png_path)
    return png_path


TBD = '<span class="tbd">マーケ部確定待ち</span>'


# ---------------------------------------------------------------------------
# 1. トップ（ホーム）
# ---------------------------------------------------------------------------
def build_top():
    blocks = [
        header_band(),
        block(1, hero("マーケティングポータル", "見出し：マーケティングポータル")),
        sec(2, "このページの目次", ["クイックリンク Web パーツ"], toc_items([
            {"label": "①4カテゴリへ", "target": "4カテゴリカード"},
            {"label": "②マーケへの業務依頼", "target": "AI-1セルフチェック→手順ガイド→業務依頼フォーム"},
            {"label": "③Webサイト流入分析", "target": "埋め込みダッシュボード"},
            {"label": "④組織体制", "target": "組織体制（最下部）"},
        ])),
        sec(3, "4カテゴリカード", ["クイックリンク Web パーツ"],
            cards([
                ("キャンディデート向け", f"1行の説明文欄<br/>{TBD}"),
                ("クライアント向け", f"1行の説明文欄<br/>{TBD}"),
                ("ブランドライブラリ", f"1行の説明文欄<br/>{TBD}"),
                ("グローバルブランドパートナーシップ", f"1行の説明文欄<br/>{TBD}"),
            ], cols=4),
            issue="T-3"),
        sec(4, "マーケへの業務依頼",
            ["①埋め込み Web パーツ", "②テキスト Web パーツ", "③行動喚起 Web パーツ"],
            f'''<div class="flowV">
              <div class="stepRow">
                <div class="stepnum">①</div>
                <div class="stepbox">{ai1box()}</div>
              </div>
              <div class="arrowDown">↓</div>
              <div class="stepRow">
                <div class="stepnum">②</div>
                <div class="stepbox guidebox">
                  <div class="t">[テキスト Web パーツ] 手順ガイド</div>
                </div>
              </div>
              <div class="arrowDown">↓</div>
              <div class="stepRow">
                <div class="stepnum">③</div>
                <div class="stepbox formbox">
                  <div class="cta-primary">[行動喚起 Web パーツ]<br/>マーケへの業務依頼フォーム</div>
                </div>
              </div>
            </div>''',
            issue="T-4"),
        sec(5, "Webサイト流入分析", ["埋め込み Web パーツ"],
            '<div class="subblock"><div class="st">埋め込みダッシュボード</div>'
            '<div class="sd">現行の表示領域をそのまま配置</div></div>'),
        sec(6, "組織体制", ["テキスト Web パーツ（H2見出し）", "画像 Web パーツ（組織図）"],
            '<div class="subblock"><div class="st">H2見出し：組織体制</div>'
            '<div class="sd">組織図画像（ページ最下部に配置）</div></div>',
            issue="T-2"),
    ]
    return render("top", 1, "1. トップ（ホーム）", blocks)


# ---------------------------------------------------------------------------
# 2. ブランドライブラリ
# ---------------------------------------------------------------------------
def build_brand():
    blocks = [
        header_band(),
        block(1, hero("ブランドライブラリ")),
        sec(2, "このページの目次", ["クイックリンク Web パーツ"], toc_items([
            {"label": "ブランドの使用について", "target": "独立2セクション（依頼前の確認／注意事項）"},
            {"label": "ガイドライン一覧", "target": "4カテゴリ見出し＋リンク箇条書き（折りたたみ格納）"},
            {"label": "ロゴ掲載申請", "target": "2ケースの手続き"},
            {"label": "ロゴ", "target": "既存ロゴ素材へのリンク"},
            {"label": "アイコン", "target": "既存アイコン素材へのリンク"},
            {"label": "フォント", "target": "フォント"},
            {"label": "画像素材", "target": "4見出し＋リンク箇条書き"},
            {"label": "各種テンプレート", "target": "5見出し＋リンク箇条書き"},
            {"label": "SNSガイドライン", "target": "リンク箇条書き"},
        ])),
        sec(3, "ブランドの使用について（1/2）：ブランドレビュー依頼の前にご確認ください",
            ["テキスト Web パーツ"],
            f'''<div class="stack">
              <div class="subblock">
                <div class="st">本文ブロック</div>
                <div class="sd">レビュー依頼の対象範囲／提出物／所要日数の目安／依頼窓口（マーケへの業務依頼フォーム）</div>
              </div>
              {ai1box()}
            </div>''',
            tone="tone-a"),
        sec(4, "ブランドの使用について（2/2）：注意事項（社外秘）",
            ["テキスト Web パーツ"],
            '<div class="subblock"><div class="st">本文ブロック</div>'
            '<div class="sd">社外への転載・二次利用の禁止事項／社外秘資料の取扱いルール</div></div>',
            tone="tone-b"),
        sec(5, "ガイドラインライブラリ", ["テキスト Web パーツ"],
            collapsible(linkgroups([
                ("全社ブランド規定", ["AKKODiS ブランドガイドライン V3（既存資料）", "AKKODiS新ビジョンガイドライン（既存資料）"]),
                ("制作物別ガイドライン", ["ビデオガイドライン（既存資料）", "フォトグラフィーガイドライン（既存資料）", "AI生成画像の使用ガイドライン（既存資料）"]),
                ("表記・用語", ["文字表記ガイドライン（既存資料）", "グローバル用語集（既存資料）", "ブランドボイスガイドライン（既存資料）"]),
                ("Akkodis Intelligenceブランド", ["ブランドガイドライン（既存資料）", "Q&A（既存資料）"]),
            ])),
            issue="B-1／B-2／B-4"),
        sec(6, "ロゴ掲載申請（手続き）",
            ["テキスト Web パーツ（番号付きリスト）", "ドキュメント ライブラリ Web パーツ（既存）"],
            cols2(
                ("①先方企業のロゴを当社サイトに掲載する場合",
                 ["AI-1でPPT/画像を事前セルフチェック", "マーケへの業務依頼フォーム提出", "Teams連絡",
                  "Salesforceサービスリクエストから押印申請", "回収", "送付", "HP掲載"],
                 "同意書テンプレ（先方ロゴ掲載用・既存ライブラリ）"),
                ("②当社社名・ロゴを他社サイト等に掲載する場合",
                 [f"フォーム起点の番号付き手順　{TBD}"],
                 "同意書テンプレ（当社ロゴ掲載用・既存ライブラリ）"),
            ),
            issue="B-3"),
        sec(7, "ロゴ", ["テキスト Web パーツ"],
            linkgroups([
                ("既存ロゴライブラリへのリンク", [
                    "AKKODiSロゴ（横組み）", "AKKODiSロゴ（縦組み）",
                    "モノクロ版", "Akkodis Intelligence版"]),
            ], single_col=True)),
        sec(8, "アイコン", ["テキスト Web パーツ"],
            linkgroups([
                ("既存アイコンライブラリへのリンク", ["UI用アイコン", "資料用アイコン", "SNS用アイコン"]),
            ], single_col=True),
            issue="B-5"),
        sec(9, "フォント", ["テキスト Web パーツ"],
            '<div class="subblock"><div class="st">本文ブロック</div>'
            '<div class="sd">現行のテキスト構成</div></div>'),
        sec(10, "画像素材", ["テキスト Web パーツ"],
            linkgroups([
                ("マーケティング部承認済み画像について", ["承認済み画像バンク（既存ライブラリ）", "利用申請はマーケへの業務依頼フォームへ"]),
                ("画像の購入について", ["推奨ストックフォトサービス一覧（既存資料）", "購入時の申請フロー（既存資料）"]),
                ("AI生成画像の使用について", ["AI生成画像の使用ガイドライン（既存資料）", "利用可否チェックリスト（既存資料）"]),
                ("使用禁止の画像について", ["使用禁止の例一覧（既存資料）", "著作権に関する注意事項（既存資料）"]),
            ])),
        sec(11, "各種テンプレート", ["テキスト Web パーツ"],
            linkgroups([
                (f"Microsoft Officeを用いた資料作成ガイドライン、テンプレート（リンク6件）　{TBD}",
                 [f"既存資料へのリンク {i}" for i in range(1, 7)]),
                (f"WEB会議用バーチャル背景（リンク1件）　{TBD}", ["既存資料へのリンク 1"]),
                (f"メール署名テンプレート（リンク1件）　{TBD}", ["既存資料へのリンク 1"]),
                (f"社員の写真掲載に関する許諾同意書（リンク1件）　{TBD}", ["既存資料へのリンク 1"]),
                (f"Formsテンプレート（リンク2件）　{TBD}", [f"既存資料へのリンク {i}" for i in range(1, 3)]),
            ], single_col=True)),
        sec(12, "SNSガイドライン", ["テキスト Web パーツ"],
            linkgroups([
                ("SNSガイドライン", [
                    "Facebook投稿におけるブランド運用ガイドライン（画像・トーン&マナー）",
                    "X（旧Twitter）投稿時のロゴ・ハッシュタグ運用ルール",
                    "LinkedIn企業ページのビジュアルガイドライン"]),
            ], single_col=True)),
        sec(13, "将来拡張（参考・本体には含めない）", [],
            '<div class="subblock"><div class="st">AI検索の高度ファセット／Copilot要約等</div>'
            '<div class="sd">将来の拡張候補として枠のみ記載</div></div>',
            tone="dashed"),
    ]
    return render("brand", 2, "2. ブランドライブラリ", blocks)


# ---------------------------------------------------------------------------
# 3. クライアント向け
# ---------------------------------------------------------------------------
def build_client():
    alt = '''<div class="altbox">
      <span class="altlabel">副案：リスト併用</span>
      <div class="t">[リスト Web パーツ]（カテゴリ／言語／更新日の列＋ビュー絞り込み）</div>
      <div class="listpart">
        <div class="filterbar">
          <span class="lbl">ビューで絞り込み:</span>
          <span class="chip">共通</span><span class="chip">ソリューション</span><span class="chip">アカデミー</span>
          <span class="chip">事例資料</span><span class="chip">お役立ち資料</span>
        </div>
        <table>
          <tr><th>タイトル（ハイパーリンク列）</th><th>カテゴリ（選択肢列）</th><th>言語（選択肢列）</th><th>更新日（任意列）</th></tr>
          <tr><td class="linkcol">会社案内（日本語版）</td><td>共通</td><td>日本語</td><td>—</td></tr>
          <tr><td class="linkcol">ソリューション紹介資料（約11本）</td><td>ソリューション</td><td>日本語／英語</td><td>—</td></tr>
          <tr><td class="linkcol">お役立ち資料（30本超）</td><td>お役立ち資料</td><td>日本語／英語</td><td>任意</td></tr>
        </table>
      </div>
    </div>'''

    blocks = [
        header_band(),
        block(1, hero("クライアント向け")),
        sec(2, "このページの目次", ["クイックリンク Web パーツ"], toc_items([
            {"label": "AKKODiS企業向けWebページ", "target": "セミナー／活用事例／Insights／サービス"},
            {"label": "公開資料ライブラリ", "target": "カテゴリ見出し＋リンク箇条書き"},
            {"label": "メルマガ配信スケジュール", "target": "配信スケジュール（社外秘）"},
        ])),
        sec(3, "AKKODiS企業向けWebページ", ["クイックリンク Web パーツ"],
            cards([
                ("セミナー", "現行ページへ"), ("活用事例", "現行ページへ"),
                ("AKKODiS Insights", "現行ページへ"), ("サービスページ", "現行ページへ"),
            ], cols=4)),
        sec(4, "公開資料ライブラリ",
            ["テキスト Web パーツ（主案）", "リスト Web パーツ（副案）"],
            f'''<div class="stack">
              {linkgroups([
                  ("共通", ["会社案内（日本語版）", "会社案内（英語版）"]),
                  ("ソリューション", ["ソリューション紹介資料（約11本・日本語／英語）"]),
                  ("アカデミー", ["アカデミー紹介資料（約8本・日本語／英語）"]),
                  (f"事例資料　{TBD}", ["事例資料へのリンク"]),
              ], single_col=True)}
              {collapsible(linkgroups([
                  (f"お役立ち資料（30本超）　{TBD}", [
                      "お役立ち資料 例①（日本語）",
                      "お役立ち資料 例②（英語）",
                      "…他 30本超"]),
              ], single_col=True))}
              {alt}
            </div>''',
            issue="C-1／C-2"),
        sec(5, "メルマガ配信スケジュール", ["クイックリンク Web パーツ"],
            '<div class="subblock"><div class="st">配信スケジュール（社外秘）</div>'
            '<div class="sd">現行の構成を配置</div></div>'),
    ]
    return render("client", 3, "3. クライアント向け", blocks)


# ---------------------------------------------------------------------------
# 4. キャンディデート向け
# ---------------------------------------------------------------------------
def build_candidate():
    blocks = [
        header_band(),
        block(1, hero("キャンディデート向けページ")),
        sec(2, "このページの目次", ["クイックリンク Web パーツ"], toc_items([
            {"label": "AKKODiS公式HP", "target": "AKKODiS People／リクルーティングムービー"},
            {"label": "マスター資料", "target": "会社説明資料（3チーム連絡先）"},
            {"label": "Social Media", "target": "X／Facebook／LinkedIn／YouTube"},
        ]), issue="CA-2"),
        sec(3, "AKKODiS公式HP",
            ["①テキスト Web パーツ", "②クイックリンク Web パーツ"],
            f'''<div class="cols2">
              <div class="col">
                <div class="t">①[テキスト Web パーツ] AKKODiS People</div>
                {linkgroups([
                    ("本文＋複数のハイパーリンク", [
                        "AKKODiS Peopleを見る（外部サイトへ）",
                        "告知投稿の依頼はマーケへの業務依頼フォームへ"]),
                ], single_col=True)}
              </div>
              <div class="col">
                <div class="t">②[クイックリンク Web パーツ] リクルーティングムービー</div>
                <div class="card" style="border:none; padding:0;">
                  <div class="icon"></div>
                  <div class="d">1タイル＝1リンク</div>
                </div>
              </div>
            </div>'''),
        sec(4, "マスター資料", ["テキスト Web パーツ"],
            '''<div class="stack">
              <div class="subblock">
                <div class="st">現行文言（そのまま掲載）</div>
                <div class="sd">「会社説明資料は各チームで管理しています。必要な場合は各チームまで直接ご連絡ください。」</div>
              </div>
              <div class="lib-meta">
                <span>中途採用向け：担当チーム連絡先</span>
                <span>新卒採用向け：担当チーム連絡先</span>
                <span>障がい者採用向け：担当チーム連絡先</span>
              </div>
            </div>''',
            issue="CA-1"),
        sec(5, "Social Media", ["クイックリンク Web パーツ"],
            cards([("X", "現行リンクへ"), ("Facebook", "現行リンクへ"),
                   ("LinkedIn", "現行リンクへ"), ("YouTube", "現行リンクへ")], cols=4)),
    ]
    return render("candidate", 4, "4. キャンディデート向け", blocks)


# ---------------------------------------------------------------------------
# 5. グローバルブランドパートナーシップ
# ---------------------------------------------------------------------------
def build_partner():
    blocks = [
        header_band(),
        block(1, hero("グローバルブランドパートナーシップ", "ヒーロー＋リード文")),
        sec(2, "このページの目次", ["クイックリンク Web パーツ"], toc_items([
            {"label": "Mercedes-AMG PETRONAS Formula One Team", "target": "該当パートナーセクション"},
            {"label": "Akkodis ASP Team", "target": "該当パートナーセクション"},
            {"label": "Stade Toulousain", "target": "該当パートナーセクション"},
        ]), issue="P-2"),
        sec(3, "各パートナーセクション（3ブロック）",
            ["テキスト Web パーツ（H2見出し）", "画像 Web パーツ", "埋め込み Web パーツ（ビデオ）"],
            '<div class="cards">' + "".join(
                f'<div class="card"><div class="t">{t}</div>'
                f'<div class="d">H2見出し（目次のアンカー先）<br/>概要／共通価値／パートナーシップビデオ</div></div>'
                for t in ["Mercedes-AMG PETRONAS<br/>Formula One Team", "Akkodis ASP Team", "Stade Toulousain"]
            ) + '</div>'),
        sec(4, "Akkodis ASP Teamの画像枠", ["画像 Web パーツ"],
            '<div class="subblock"><div class="st">画像枠</div>'
            '<div class="sd">他の2チームと同じ構造で配置</div></div>',
            issue="P-1"),
    ]
    return render("partner", 5, "5. グローバルブランドパートナーシップ", blocks)


# ---------------------------------------------------------------------------
# コメント（画像から外した説明文の清書版）
# ---------------------------------------------------------------------------
HEADER_COMMENT = (
    "サイト内検索はサイトヘッダーの標準機能を使い、ページ本文には検索ボックスを配置しない。"
    "全ページ共通でサイト共通エージェント（AI-2）をヘッダーから起動でき、利用者本人のアクセス権の範囲で"
    "出典リンク付きの回答を返す（回答の根拠になるのはページ本文とドキュメントライブラリで、"
    "リストのデータは対象外とMicrosoft公式ドキュメントに記載されている）。"
    "利用にはMicrosoft 365 Copilotライセンスまたは従量課金の設定が必要なため、テナント管理者に可否を確認する。"
)

COMMENTS = {
    "top": {
        "page_title": "トップ（ホーム）",
        "sections": [
            {"no": 0, "title": "共通ヘッダー（サイト内検索／AI-2）",
             "part": "サイトヘッダー（システム標準）／SharePoint agent",
             "issue": "", "comment": HEADER_COMMENT},
            {"no": 1, "title": "ヒーロー", "part": "ヒーロー Web パーツ", "issue": "",
             "comment": "見出し「マーケティングポータル」を掲出する現行のヒーローをそのまま使う。"},
            {"no": 2, "title": "このページの目次", "part": "クイックリンク Web パーツ", "issue": "",
             "comment": "目次専用のWebパーツは提供されていないため、テキストWebパーツの見出しから自動生成される"
                        "ページ内アンカーを、クイックリンクWebパーツで束ねて目次ブロックとして表現する。"
                        "この目次パターンは全ページで共通とする。"},
            {"no": 3, "title": "4カテゴリカード", "part": "クイックリンク Web パーツ", "issue": "T-3",
             "comment": "各カードの直下に1行の説明文欄を設け、カードだけでは伝わらない遷移先の内容を補う。"
                        "クイックリンクWebパーツはレスポンシブに4件を横並び表示できるため4列で配置する"
                        "（3列までという制約はセクションの列レイアウトに適用されるもので、このパーツ自体には掛からない）。"
                        "説明文の文面はマーケティング部で確定する。"},
            {"no": 4, "title": "マーケへの業務依頼",
             "part": "埋め込み Web パーツ／テキスト Web パーツ／行動喚起 Web パーツ", "issue": "T-4",
             "comment": "①AI-1セルフチェック → ②手順ガイド → ③業務依頼フォーム を上から下へ全幅で並べる縦積みの導線とし、"
                        "正式な依頼窓口はフォーム1つに統一する。①のAI-1（ブランドセルフチェッカー）はCopilot Studioで作成した"
                        "エージェントを埋め込みWebパーツで設置し、アップロードした資料をブランドガイドラインと照合して"
                        "一次スクリーニングの気づきを返す（色やロゴの厳密な判定は保証せず、最終承認は人が行う）。"
                        "Copilot Studioのライセンス、PowerPointファイルのアップロード機能、認証なしのWebサイトチャネル公開の可否は、"
                        "テナント管理者と実機検証で確認する。"},
            {"no": 5, "title": "Webサイト流入分析", "part": "埋め込み Web パーツ", "issue": "",
             "comment": "現行の埋め込みダッシュボードをそのまま維持する。PC表示を対象とし、端末別の出し分けは行わない。"},
            {"no": 6, "title": "組織体制", "part": "テキスト Web パーツ（H2見出し）／画像 Web パーツ", "issue": "T-2",
             "comment": "見出し「組織体制」をH2で置き、ページ内アンカーの生成対象にする。"
                        "配置はページ最下部のままとし、目次からの到達手段だけを追加する。"},
        ],
    },
    "brand": {
        "page_title": "ブランドライブラリ",
        "sections": [
            {"no": 0, "title": "共通ヘッダー（サイト内検索／AI-2）",
             "part": "サイトヘッダー（システム標準）／SharePoint agent",
             "issue": "", "comment": HEADER_COMMENT},
            {"no": 1, "title": "ヒーロー", "part": "ヒーロー Web パーツ", "issue": "",
             "comment": "現行のヒーローをそのまま使う。"},
            {"no": 2, "title": "このページの目次", "part": "クイックリンク Web パーツ", "issue": "",
             "comment": "9項目のページ内アンカーを目次ブロックとして束ねる。"
                        "ページ上部の目次と役割が重なるアンカーカードは別途設けない。"},
            {"no": 3, "title": "ブランドの使用について（1/2）：ブランドレビュー依頼の前にご確認ください",
             "part": "テキスト Web パーツ", "issue": "",
             "comment": "レビュー依頼の対象範囲・提出物・所要日数の目安・依頼窓口を本文として記載する。"
                        "セクションの背景色はセクション単位でのみ設定でき、テキストWebパーツ内の段落ごとに色を付ける機能は"
                        "Microsoft公式ドキュメントに記載がないため、案内トーンと注意喚起トーンの色分けはセクションを分けて実現する。"
                        "AI-1（ブランドセルフチェッカー）はこのセクションに設置し、ロゴ掲載申請のセクションからは同じ導線を参照して"
                        "二重に設置しない。"},
            {"no": 4, "title": "ブランドの使用について（2/2）：注意事項（社外秘）",
             "part": "テキスト Web パーツ", "issue": "",
             "comment": "社外への転載・二次利用の禁止事項と、社外秘資料の取扱いルールを本文として記載する。"
                        "案内トーンのセクションと区別するため、独立したセクションとして背景色を変える。"
                        "本文中の黄色マーカー装飾は使わない。"},
            {"no": 5, "title": "ガイドラインライブラリ", "part": "テキスト Web パーツ", "issue": "B-1／B-2／B-4",
             "comment": "新規のドキュメントライブラリは作らず、目的軸の4カテゴリ見出しと既存資料へのリンク箇条書きを"
                        "テキストWebパーツで表現する。情報量が多いためページは分割せず、「このセクションを折りたたみ可能にする」を"
                        "有効にして既定は折りたたみ状態で格納する。リンクラベルの表記は「AKKODiS」に統一し、"
                        "掲載前に既存ガイドラインの重複・旧版の棚卸しを運用作業として実施する。"},
            {"no": 6, "title": "ロゴ掲載申請（手続き）",
             "part": "テキスト Web パーツ／ドキュメント ライブラリ Web パーツ", "issue": "B-3",
             "comment": "先方企業のロゴを当社サイトへ掲載する場合と、当社の社名・ロゴを他社サイト等へ掲載する場合の2ケースを、"
                        "フォームを起点とする番号付き手順として並べる。2カラムはセクション内の2列レイアウトで実現する。"
                        "②の具体的な手順はマーケティング部が現物を確認して確定するため、ここでは構造要件のみを定める。"},
            {"no": 7, "title": "ロゴ", "part": "テキスト Web パーツ", "issue": "",
             "comment": "ロゴ素材そのものは新規ライブラリを作らず、既存のロゴ素材ライブラリへのリンクをテキストWebパーツに並べる。"
                        "手続きを扱う「ロゴ掲載申請」とは別のセクションである旨を本文に明記する。"
                        "既存ライブラリの中身を一覧表示したい場合は、既存ライブラリを参照するドキュメント ライブラリ Web パーツを"
                        "代替手段として使える。"},
            {"no": 8, "title": "アイコン", "part": "テキスト Web パーツ", "issue": "B-5",
             "comment": "アイコン素材も新規ライブラリを作らず、既存のアイコン素材ライブラリへのリンクを並べる。"
                        "ここで扱うのは素材データであり、画面のナビゲーション用アイコンとは別物である旨を本文に明記する。"},
            {"no": 9, "title": "フォント", "part": "テキスト Web パーツ", "issue": "",
             "comment": "現行のテキスト構成を維持し、目次から到達できるようにする。"},
            {"no": 10, "title": "画像素材", "part": "テキスト Web パーツ", "issue": "",
             "comment": "1つのテキストWebパーツの中に4つの見出しと各リンクの箇条書きを置き、クイックリンクWebパーツは併用しない。"
                        "図中の2列表示はワイヤーフレーム上の便宜的な配置であり、セクションの列レイアウトを指定するものではない。"},
            {"no": 11, "title": "各種テンプレート", "part": "テキスト Web パーツ", "issue": "",
             "comment": "5区分の見出しと実際のリンク件数（6／1／1／1／2件）を、テキストWebパーツ1つの中に見出し＋箇条書きで表現する。"
                        "強調表示されたコンテンツWebパーツはサイトやライブラリ単位の動的クエリでファイル・ページを一覧表示するパーツで、"
                        "個別のリンクを1件ずつ指定する少数のリンク集には使用しない。掲載する個々のリンク先はマーケティング部が確定する。"},
            {"no": 12, "title": "SNSガイドライン", "part": "テキスト Web パーツ", "issue": "",
             "comment": "3件のリンクは見出しが長いため、クイックリンクWebパーツではなくテキストWebパーツのリンク付き箇条書きで表現する。"
                        "クイックリンクの見出し文字数の上限は公開情報に数値の記載がないため、実機検証で確認する。"},
            {"no": 13, "title": "将来拡張（参考・本体には含めない）", "part": "—", "issue": "",
             "comment": "ファセット検索やCopilotによる要約などの高度な絞り込みは、本構築の範囲には含めない。"
                        "カスタム開発の許可を前提とする将来の拡張候補として、枠の記載にとどめる。"},
        ],
    },
    "client": {
        "page_title": "クライアント向け",
        "sections": [
            {"no": 0, "title": "共通ヘッダー（サイト内検索／AI-2）",
             "part": "サイトヘッダー（システム標準）／SharePoint agent",
             "issue": "", "comment": HEADER_COMMENT},
            {"no": 1, "title": "ヒーロー", "part": "ヒーロー Web パーツ", "issue": "",
             "comment": "現行のヒーローをそのまま使う。"},
            {"no": 2, "title": "このページの目次", "part": "クイックリンク Web パーツ", "issue": "",
             "comment": "3項目のページ内アンカーを目次ブロックとして束ねる。"
                        "遷移先が特定できない「資料を探す」という項目は設けない"
                        "（本文検索はサイトヘッダーの標準検索が担い、資料への導線は「公開資料ライブラリ」と重複するため）。"},
            {"no": 3, "title": "AKKODiS企業向けWebページ", "part": "クイックリンク Web パーツ", "issue": "",
             "comment": "セミナー・活用事例・AKKODiS Insights・サービスページの4つを現行の内容のまま並べる。"
                        "クイックリンクWebパーツはレスポンシブに4件を横並び表示できるため4列で配置する。"},
            {"no": 4, "title": "公開資料ライブラリ",
             "part": "テキスト Web パーツ（主案）／リスト Web パーツ（副案）", "issue": "C-1／C-2",
             "comment": "主案は、カテゴリ見出しと資料リンクの箇条書きをテキストWebパーツで表現し（掲載するのは資料への"
                        "リンクでありファイル実体ではない。更新日の記載は任意）、30本を超えるお役立ち資料は折りたたみ"
                        "セクションに格納する。SharePoint agentが回答の根拠に使えるのはページ本文とドキュメントライブラリで、"
                        "リストのデータは対象外だとMicrosoft公式ドキュメントに明記されているため、資料リンクはページ本文に置く。"
                        "点線枠の副案は、カテゴリ・言語・更新日の列とビューによる絞り込みを人が使うための補助としてリスト"
                        "Webパーツを併設する案で、主案が持たない動的な絞り込みを補える一方、エージェントの回答根拠にはならない。"},
            {"no": 5, "title": "メルマガ配信スケジュール", "part": "クイックリンク Web パーツ", "issue": "",
             "comment": "現行の構成をそのまま維持し、社外秘である旨の表記も残す。"},
        ],
    },
    "candidate": {
        "page_title": "キャンディデート向け",
        "sections": [
            {"no": 0, "title": "共通ヘッダー（サイト内検索／AI-2）",
             "part": "サイトヘッダー（システム標準）／SharePoint agent",
             "issue": "", "comment": HEADER_COMMENT},
            {"no": 1, "title": "ヒーロー", "part": "ヒーロー Web パーツ", "issue": "",
             "comment": "現行のヒーローをそのまま使う。"},
            {"no": 2, "title": "このページの目次", "part": "クイックリンク Web パーツ", "issue": "CA-2",
             "comment": "3項目のページ内アンカーを目次ブロックとして束ね、3セクション構成を維持する。"},
            {"no": 3, "title": "AKKODiS公式HP",
             "part": "テキスト Web パーツ／クイックリンク Web パーツ", "issue": "",
             "comment": "「AKKODiS People」は、説明文と複数のハイパーリンク（採用サイトへのリンク、"
                        "告知投稿の依頼はマーケへの業務依頼フォームへ）をテキストWebパーツで表現する。"
                        "クイックリンクWebパーツは1タイルにつき1リンクという構成で、1つのタイルに副リンクを追加する設定項目が"
                        "存在しないため、副リンクが必要なブロックにはテキストWebパーツを使う。"
                        "「リクルーティングムービー」は1タイル1リンクのクイックリンクWebパーツのまま維持する。"},
            {"no": 4, "title": "マスター資料", "part": "テキスト Web パーツ", "issue": "CA-1",
             "comment": "会社説明資料は各チームが管理しているため、現行の案内文をそのままテキストWebパーツに掲載し、"
                        "資料への直リンク化は行わない。中途採用・新卒採用・障がい者採用の3チームの連絡先を現行どおり併記する。"},
            {"no": 5, "title": "Social Media", "part": "クイックリンク Web パーツ", "issue": "",
             "comment": "X・Facebook・LinkedIn・YouTubeの4タイルを横並びで配置する。"},
        ],
    },
    "partner": {
        "page_title": "グローバルブランドパートナーシップ",
        "sections": [
            {"no": 0, "title": "共通ヘッダー（サイト内検索／AI-2）",
             "part": "サイトヘッダー（システム標準）／SharePoint agent",
             "issue": "", "comment": HEADER_COMMENT},
            {"no": 1, "title": "ヒーロー", "part": "ヒーロー Web パーツ", "issue": "",
             "comment": "ヒーローとリード文を現行のまま維持する。"},
            {"no": 2, "title": "このページの目次", "part": "クイックリンク Web パーツ", "issue": "P-2",
             "comment": "3パートナーへのページ内アンカーを目次ブロックとして束ね、全ページ共通の目次パターンに合わせる。"
                        "リンク先が目次と同じ3セクションになるサムネイルカードは、役割が重複するため設けない。"},
            {"no": 3, "title": "各パートナーセクション（3ブロック）",
             "part": "テキスト Web パーツ／画像 Web パーツ／埋め込み Web パーツ", "issue": "",
             "comment": "Mercedes-AMG PETRONAS Formula One Team／Akkodis ASP Team／Stade Toulousainの3ブロックを、"
                        "概要・共通価値・パートナーシップビデオという現行の構成のまま並べる。"
                        "各ブロック先頭のH2見出しが目次のアンカー先になる。"},
            {"no": 4, "title": "Akkodis ASP Teamの画像枠", "part": "画像 Web パーツ", "issue": "P-1",
             "comment": "Akkodis ASP Teamの画像枠を、他の2チームと同じ構造に揃える。"},
        ],
    },
}


def build_comments():
    """コメント本文に、render() で実測したセクションの縦位置 y / 高さ h を合成して書き出す。

    y / h は出力PNGの実ピクセル基準。Excel でWF画像の真横にコメント行を並べるための座標。
    """
    path = os.path.join(OUT, "comments.json")
    out = {}
    for page, data in COMMENTS.items():
        geom = GEOM.get(page, {})
        img = geom.get("__img__")
        page_out = {"page_title": data["page_title"]}
        if img:
            page_out["image_width"] = img["w"]
            page_out["image_height"] = img["h"]
        secs = []
        for s in data["sections"]:
            g = geom.get(str(s["no"]))
            item = dict(s)
            if g:
                item["y"] = g["y"]
                item["h"] = g["h"]
            secs.append(item)
        page_out["sections"] = secs
        out[page] = page_out

    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print("saved", path)

    # --- 検証 -------------------------------------------------------------
    for page, data in out.items():
        ys = [s.get("y") for s in data["sections"]]
        assert all(v is not None for v in ys), f"{page}: 座標の取れないセクションがある"
        assert ys == sorted(ys), f"{page}: y が番号順に単調増加していない -> {ys}"
        last = data["sections"][-1]
        assert last["y"] + last["h"] <= data["image_height"], (
            f"{page}: 最終セクションが画像高さを超えている "
            f"({last['y'] + last['h']} > {data['image_height']})"
        )
        print(f"  OK {page}: {len(ys)} sections, image_height={data['image_height']}, "
              f"last bottom={last['y'] + last['h']}")
    return path


if __name__ == "__main__":
    build_top()
    build_brand()
    build_client()
    build_candidate()
    build_partner()
    build_comments()
    print("all done")

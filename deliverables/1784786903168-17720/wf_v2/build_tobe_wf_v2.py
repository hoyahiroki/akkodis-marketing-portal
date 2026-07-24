#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ToBe ワイヤーフレーム v2 生成スクリプト（判読性最優先・大判レンダリング）
- 入力: tobe-wf-spec-v2.md（設計スペック）に忠実に、全5ページ・6枚のWF画像を生成
- HTML/CSS を Playwright(Chromium) でPNG化。配色・装飾は作り込まず灰色矩形+ラベルのみ（designer-uiが後工程）
- 検索ボックスは本文に描かない／目次は全ページ共通パターン／将来拡張は点線枠で区別
- 前回(v1)の反省: 「小さすぎ・雑」→ 本v2は幅1120px・本文14px以上・余白を広くとって判読性を最優先する
- 出力後、PIL で6枚のPNGを1つのPDF（ToBe_WF_v2_承認用.pdf）にまとめる（各PNG冒頭にページ名見出しを内包）
"""
import os
from playwright.sync_api import sync_playwright
from PIL import Image

OUT = "/workspace/akkodis-marketing-portal/deliverables/1784786903168-17720/wf_v2"
os.makedirs(OUT, exist_ok=True)

CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

WIDTH = 1120  # 本文幅（要件: 1000〜1200px程度）

# ---------------------------------------------------------------------------
# CSS（判読性優先: 本文14px以上、見出し17-22px、余白広め）
# ---------------------------------------------------------------------------
CSS = """
* { box-sizing: border-box; margin:0; padding:0; }
body {
  font-family: "IPAGothic","IPAPGothic","Noto Sans CJK JP",sans-serif;
  background:#e7e9ec; padding:0 0 40px 0; width: __WIDTH__px;
}
.wrap { width:100%; }

/* ページ名見出し（PDF化時にも機能する常設タイトルバー） */
.titlebar {
  background:#1c2733; color:#fff; padding:18px 26px;
  display:flex; justify-content:space-between; align-items:baseline;
}
.titlebar .name { font-size:24px; font-weight:bold; letter-spacing:.5px; }
.titlebar .sub { font-size:13px; color:#b8c2cf; }

/* 検索帯（システム標準・本文外・R1） */
.searchband {
  background:#454e59; color:#e9ecf0; padding:12px 26px;
  display:flex; align-items:center; gap:14px; font-size:13px;
  border-bottom:3px solid #1c2733;
}
.searchband .tag {
  background:#6c7480; border:1px solid #939ba6; color:#fff;
  padding:3px 10px; border-radius:3px; font-size:11.5px; font-weight:bold;
}
.searchband .desc { color:#cfd4da; font-size:12.5px; }

.page { background:#fff; border:2px solid #333; padding:0 0 6px 0; }

/* ヒーロー */
.hero {
  margin:20px 24px; min-height:78px; border:2px solid #7a7a7a;
  background:repeating-linear-gradient(45deg,#dfe1e3,#dfe1e3 12px,#d0d2d5 12px,#d0d2d5 24px);
  display:flex; flex-direction:column; justify-content:center; padding:14px 22px;
}
.hero .t { font-size:20px; font-weight:bold; color:#20242a; }
.hero .s { font-size:13.5px; color:#4a4f57; margin-top:5px; }
.part-tag {
  display:inline-block; font-family:"Menlo","Consolas",monospace; font-size:12.5px;
  background:#e2e4e8; border:1px solid #9a9fa6; color:#33383e;
  padding:3px 10px; border-radius:4px; margin-right:8px;
}

/* 目次（全ページ共通パターン G-1） */
.toc {
  margin:22px 24px; border:2.5px solid #2f4f9e; background:#eef2fb; padding:16px 20px;
}
.toc .head {
  font-size:17px; font-weight:bold; color:#1c3570; margin-bottom:4px;
  display:flex; align-items:center; gap:10px;
}
.toc .headnote { font-size:11.5px; color:#5567a0; margin-bottom:12px; }
.toc-items { display:flex; flex-wrap:wrap; gap:10px; }
.toc-item {
  background:#fff; border:1.5px solid #3b5bb0; border-radius:20px;
  padding:9px 16px; font-size:14px; color:#1c3570; line-height:1.5;
}
.toc-item b { color:#0d47a1; }
.toc-item .new { background:#ffe9c2; border:1px solid #c58a2e; color:#7a4e0d; font-size:10.5px; padding:1px 6px; border-radius:3px; margin-left:6px; }
.toc-item.removed { text-decoration:line-through; color:#999; border-color:#bbb; background:#f3f3f3; }

/* 判断根拠ノート帯 */
.rationale {
  margin:14px 24px 22px; border:2px solid #7a7a7a; background:#f4f4f5; padding:12px 18px;
  font-size:13px; color:#3a3a3a; line-height:1.8;
}
.rationale .h { font-size:13.5px; font-weight:bold; color:#222; margin-bottom:4px; }

/* セクションブロック本体 */
.sec {
  margin:20px 24px; border:2px solid #7a7a7a; background:#f7f7f8; padding:18px 22px;
}
.sec.new { border-color:#2f6b3f; background:#f2f8f3; }
.sec .sechead { display:flex; justify-content:space-between; align-items:flex-start; gap:14px; margin-bottom:10px; }
.sec .num { font-family:monospace; font-size:12px; color:#777; }
.sec .title { font-size:18px; font-weight:bold; color:#181b1f; margin:6px 0 8px; }
.sec .badges { display:flex; gap:6px; flex-wrap:wrap; }
.badge { font-size:11px; padding:2px 8px; border-radius:3px; border:1px solid #999; color:#444; background:#eee; white-space:nowrap;}
.badge.rule { background:#e7edf9; border-color:#7d94c9; color:#28407e; }
.badge.newf { background:#e3f3e6; border-color:#4f9463; color:#1d5c30; }
.badge.ok { background:#e6f0fb; border-color:#5c85b8; color:#204a78; }
.desc { font-size:14.5px; line-height:1.85; color:#2b2b2b; }
.note {
  margin-top:10px; font-size:13px; color:#555; line-height:1.75;
  border-left:3px solid #9aa; padding-left:10px;
}
.tbd {
  color:#8a4b00; background:#fff3de; border:1px solid #d69a3f;
  padding:1px 7px; border-radius:3px; font-size:13px; font-weight:bold;
}

/* 主導線/副導線ボタン風 */
.cta-primary {
  display:inline-block; background:#2b3a55; color:#fff; font-size:16px; font-weight:bold;
  padding:14px 26px; border-radius:5px; border:2px solid #1c2733;
}
.cta-secondary {
  display:inline-block; margin-top:12px; margin-left:26px; font-size:13.5px; color:#2b3a55;
  border-bottom:1.5px solid #7f8aa0; padding-bottom:2px;
}
.cta-relation { font-size:13px; color:#555; margin:12px 0 4px 4px; line-height:1.8; }
.nest-wrap { border-left:4px solid #cfd4da; padding-left:18px; margin-top:6px; }

/* カード群 */
.cards { display:flex; gap:14px; flex-wrap:wrap; margin-top:6px; }
.card {
  flex:1 1 220px; border:2px solid #8a8f96; background:#fff; padding:14px 14px;
}
.card .icon { width:34px; height:34px; background:#c9cbce; border-radius:6px; margin-bottom:8px; }
.card .t { font-size:14.5px; font-weight:bold; color:#1e2126; margin-bottom:6px; }
.card .d { font-size:12.5px; color:#555; line-height:1.6; }

/* 2カラム */
.cols2 { display:flex; gap:18px; margin-top:8px; }
.col { flex:1; border:1.5px solid #9a9fa6; background:#fff; padding:14px 16px; }
.col .t { font-size:14.5px; font-weight:bold; margin-bottom:8px; color:#1e2126; }
.col ol { margin-left:18px; font-size:13px; line-height:2; color:#2b2b2b; }
.col .link { margin-top:10px; font-size:12px; color:#2b3a55; border-top:1px dashed #aaa; padding-top:8px; }

/* ライブラリ（メタ列＋カテゴリカード） */
.lib-meta { display:flex; gap:8px; flex-wrap:wrap; margin:10px 0 12px; }
.lib-meta span {
  font-size:11.5px; background:#e2e4e8; border:1px solid #9a9fa6; color:#333;
  padding:3px 9px; border-radius:3px;
}
.lib-cats { display:flex; gap:12px; flex-wrap:wrap; }
.lib-cat { flex:1 1 220px; border:1.5px solid #8a8f96; background:#fbfbfc; padding:12px 14px; }
.lib-cat .n { font-size:13.5px; font-weight:bold; margin-bottom:6px; color:#1e2126; }
.lib-cat ul { list-style:none; }
.lib-cat li { font-size:12px; color:#4a4a4a; padding:3px 0 3px 14px; position:relative; line-height:1.6; }
.lib-cat li:before { content:"▪"; position:absolute; left:0; color:#999; font-size:8px; top:8px; }

/* 折りたたみ/ハブ&スポーク 副案 */
.altbox {
  margin-top:14px; border:2px dashed #6c7480; background:#f1f2f4; padding:12px 16px;
}
.altbox .tag { display:inline-block; font-size:11px; background:#dfe1e5; border:1px solid #8a8f96; color:#444; padding:2px 8px; border-radius:3px; margin-bottom:6px;}
.altbox .body { font-size:12.5px; color:#4a4a4a; line-height:1.75; }

/* 将来拡張（点線枠・L2/L3） */
.future {
  margin:20px 24px; border:2.5px dashed #b5762a; background:#fbf3e8; padding:16px 20px;
}
.future .tag {
  display:inline-block; font-size:12px; font-weight:bold; color:#8a5518; background:#f3dfc0;
  border:1px solid #b5762a; padding:3px 10px; border-radius:4px; margin-bottom:8px;
}
.future .t { font-size:14.5px; font-weight:bold; color:#6f4413; margin-bottom:6px; }
.future .d { font-size:13px; color:#6f5535; line-height:1.75; }

/* 比較表（役割A/B等） */
.cmp { margin:16px 24px 4px; border:2px solid #7a7a7a; }
.cmp table { width:100%; border-collapse:collapse; }
.cmp th, .cmp td { border:1px solid #b3b6ba; padding:9px 12px; font-size:12.5px; text-align:left; vertical-align:top; line-height:1.6; }
.cmp th { background:#e9ebee; font-size:12.5px; }
.cmp caption { text-align:left; font-size:14px; font-weight:bold; padding:10px 12px; background:#f0f1f3; caption-side:top; border:1px solid #b3b6ba; border-bottom:none;}

/* 変更点まとめ（下部） */
.changes {
  margin:26px 24px 20px; border:2.5px solid #2f6b3f; background:#eef7ef; padding:18px 22px;
}
.changes .h { font-size:15.5px; font-weight:bold; color:#1c4a28; margin-bottom:10px; display:flex; align-items:center; gap:8px;}
.changes .h .pin { background:#2f6b3f; color:#fff; font-size:11px; padding:2px 8px; border-radius:3px; }
.changes ul { margin-left:20px; }
.changes li { font-size:13.5px; color:#274d30; line-height:1.9; }

.footnote { text-align:center; font-size:11.5px; color:#8a8f96; padding:14px 10px 4px; }
"""

HTML_TMPL = """<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head>
<body><div class="wrap">{body}</div></body></html>"""


# ---------------------------------------------------------------------------
# ビルダー関数群
# ---------------------------------------------------------------------------
def titlebar(name, idx, total, sub=""):
    return f'''<div class="titlebar">
      <div class="name">ToBe ワイヤーフレーム v2　｜　{name}　（{idx}/{total}）</div>
      <div class="sub">{sub or "AKKODiSマーケティングポータル・承認用WF"}</div>
    </div>'''


def searchband():
    return '''<div class="searchband">
      <span class="tag">システム標準・編集不可</span>
      <span>サイトヘッダー：Microsoft Search（既定表示・位置固定・全ページ共通）</span>
      <span class="desc">※本文コンポーネントではない。ページ本文には検索ボックスを配置しない（R1）</span>
    </div>'''


def hero(title, sub="現状維持"):
    return f'<div class="hero"><div class="t">[Hero web part] {title}</div><div class="s">{sub}</div></div>'


def toc(items, note='※専用の目次web partは存在しない。テキストweb partの見出し（アンカー自動生成）を、クイックリンクweb partで束ねた「目次ブロック」として表現（R4・全ページ共通パターン）'):
    its = ""
    for it in items:
        cls = "toc-item"
        badge = ""
        label = it["label"]
        target = it.get("target", "")
        if it.get("removed"):
            cls += " removed"
            its += f'<div class="{cls}">{label}（削除）</div>'
            continue
        if it.get("new"):
            badge = '<span class="new">新規</span>'
        its += f'<div class="{cls}">→ <b>{label}</b>{badge}　｜遷移先: {target}</div>'
    return f'''<div class="toc">
      <div class="head">[クイックリンク web part] このページの目次</div>
      <div class="headnote">{note}</div>
      <div class="toc-items">{its}</div>
    </div>'''


def rationale(head, body):
    return f'<div class="rationale"><div class="h">{head}</div>{body}</div>'


def sec(num, title, part, desc, note=None, badges=None, new=False, extra=""):
    badges_html = "".join(f'<span class="badge {b[1]}">{b[0]}</span>' for b in (badges or []))
    note_html = f'<div class="note">{note}</div>' if note else ""
    klass = "sec new" if new else "sec"
    return f'''<div class="{klass}">
      <div class="sechead">
        <div>
          <div class="num">SECTION {num}</div>
          <div><span class="part-tag">{part}</span>{'<span class="badge newf">新設</span>' if new else ''}</div>
          <div class="title">{title}</div>
        </div>
        <div class="badges">{badges_html}</div>
      </div>
      <div class="desc">{desc}</div>
      {extra}
      {note_html}
    </div>'''


def cta_block(num, title, part, primary_label, secondary_label, relation, note=None, badges=None):
    badges_html = "".join(f'<span class="badge {b[1]}">{b[0]}</span>' for b in (badges or []))
    note_html = f'<div class="note">{note}</div>' if note else ""
    return f'''<div class="sec">
      <div class="sechead">
        <div>
          <div class="num">SECTION {num}</div>
          <div><span class="part-tag">{part}</span></div>
          <div class="title">{title}</div>
        </div>
        <div class="badges">{badges_html}</div>
      </div>
      <div class="cta-primary">[Call to action web part]　{primary_label}</div>
      <div class="cta-relation">{relation}</div>
      <div class="nest-wrap"><div class="cta-secondary">[テキスト web part] {secondary_label}</div></div>
      {note_html}
    </div>'''


def cards_block(num, title, part, cards, note=None, badges=None, new=False):
    badges_html = "".join(f'<span class="badge {b[1]}">{b[0]}</span>' for b in (badges or []))
    cs = "".join(
        f'<div class="card"><div class="icon"></div><div class="t">{c[0]}</div><div class="d">{c[1]}</div></div>'
        for c in cards
    )
    note_html = f'<div class="note">{note}</div>' if note else ""
    klass = "sec new" if new else "sec"
    return f'''<div class="{klass}">
      <div class="sechead">
        <div>
          <div class="num">SECTION {num}</div>
          <div><span class="part-tag">{part}</span></div>
          <div class="title">{title}</div>
        </div>
        <div class="badges">{badges_html}</div>
      </div>
      <div class="cards">{cs}</div>
      {note_html}
    </div>'''


def cols2_block(num, title, part, col1, col2, note=None, badges=None):
    """col1/col2 = (title, [steps], link_text or None)"""
    badges_html = "".join(f'<span class="badge {b[1]}">{b[0]}</span>' for b in (badges or []))
    def render_col(c):
        lis = "".join(f"<li>{s}</li>" for s in c[1])
        link = f'<div class="link">🔗 [ドキュメント ライブラリ web part] {c[2]}</div>' if c[2] else ""
        return f'<div class="col"><div class="t">{c[0]}</div><ol>{lis}</ol>{link}</div>'
    note_html = f'<div class="note">{note}</div>' if note else ""
    return f'''<div class="sec">
      <div class="sechead">
        <div>
          <div class="num">SECTION {num}</div>
          <div><span class="part-tag">{part}</span></div>
          <div class="title">{title}</div>
        </div>
        <div class="badges">{badges_html}</div>
      </div>
      <div class="cols2">{render_col(col1)}{render_col(col2)}</div>
      {note_html}
    </div>'''


def library_block(num, title, part, meta, cats, note=None, badges=None, altbox=None, new=False):
    badges_html = "".join(f'<span class="badge {b[1]}">{b[0]}</span>' for b in (badges or []))
    metas = "".join(f"<span>{m}</span>" for m in meta)
    cathtml = "".join(
        f'<div class="lib-cat"><div class="n">{c[0]}</div><ul>{"".join(f"<li>{x}</li>" for x in c[1])}</ul></div>'
        for c in cats
    )
    note_html = f'<div class="note">{note}</div>' if note else ""
    alt_html = ""
    if altbox:
        alt_html = f'<div class="altbox"><span class="tag">副案（点線＝要テナント確認／要社長判断）</span><div class="body">{altbox}</div></div>'
    klass = "sec new" if new else "sec"
    return f'''<div class="{klass}">
      <div class="sechead">
        <div>
          <div class="num">SECTION {num}</div>
          <div><span class="part-tag">{part}</span></div>
          <div class="title">{title}</div>
        </div>
        <div class="badges">{badges_html}</div>
      </div>
      <div class="lib-meta">{metas}</div>
      <div class="lib-cats">{cathtml}</div>
      {alt_html}
      {note_html}
    </div>'''


def future_block(title, desc):
    return f'''<div class="future">
      <span class="tag">将来拡張（L2/L3・本体には含めない）</span>
      <div class="t">{title}</div>
      <div class="d">{desc}</div>
    </div>'''


def cmp_table(caption, rows):
    trs = "".join(f"<tr><th>{r[0]}</th><td>{r[1]}</td><td>{r[2]}</td></tr>" for r in rows)
    return f'''<div class="cmp"><table>
      <caption>{caption}</caption>
      <tr><th></th><th>役割A：3パートナーのサムネイルカード</th><th>役割B：ページ内目次</th></tr>
      {trs}
    </table></div>'''


def changes_box(lines):
    lis = "".join(f"<li>{l}</li>" for l in lines)
    return f'''<div class="changes">
      <div class="h"><span class="pin">変更点</span>このページの変更点（社長指摘の反映）</div>
      <ul>{lis}</ul>
    </div>'''


def footnote(text):
    return f'<div class="footnote">{text}</div>'


def render(name, idx, total, page_label, blocks, sub=""):
    css = CSS.replace("__WIDTH__", str(WIDTH))
    body = titlebar(page_label, idx, total, sub) + searchband() + '<div class="page">' + "".join(blocks) + '</div>'
    html = HTML_TMPL.format(css=css, body=body)
    html_path = os.path.join(OUT, f"_tmp_{name}.html")
    with open(html_path, "w") as f:
        f.write(html)
    png_path = os.path.join(OUT, f"{name}.png")
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME, args=["--no-sandbox"])
        page = b.new_page(viewport={"width": WIDTH, "height": 900}, device_scale_factor=2)
        page.goto("file://" + html_path)
        page.wait_for_timeout(100)
        page.screenshot(path=png_path, full_page=True)
        b.close()
    os.remove(html_path)
    print("saved", png_path)
    return png_path


# ---------------------------------------------------------------------------
# 1. トップ（ホーム）
# ---------------------------------------------------------------------------
def build_top():
    blocks = [
        hero("マーケティングポータル", "見出し「マーケティングポータル」（現状維持）"),
        toc([
            {"label": "①4カテゴリへ", "target": "4カテゴリカード"},
            {"label": "②マーケへの業務依頼", "target": "マーケへの業務依頼（フォーム＋手順ガイド）"},
            {"label": "③Webサイト流入分析", "target": "埋め込みダッシュボード"},
            {"label": "④組織体制", "target": "組織体制（最下部・配置維持）"},
        ]),
        cards_block(
            3, "4カテゴリカード", "[クイックリンク web part]",
            [
                ("キャンディデート向け", '<span class="tbd">（マーケ部確定待ち・例）</span>採用検討中の方向けページへの案内文'),
                ("クライアント向け", '<span class="tbd">（マーケ部確定待ち・例）</span>クライアント向け公開資料・案内文'),
                ("ブランドライブラリ", '<span class="tbd">（マーケ部確定待ち・例）</span>ロゴ・ガイドライン等の素材案内文'),
                ("グローバルブランドパートナーシップ", '<span class="tbd">（マーケ部確定待ち・例）</span>提携パートナー紹介の案内文'),
            ],
            note="各カード直下に1行の説明文欄を新設（T-3対応）。文言はマーケ部確定待ちのプレースホルダーとして描き、創作文言を確定情報のように見せない。",
            badges=[("T-3", "ok")],
        ),
        cta_block(
            4, "マーケへの業務依頼（再設計・核）", "[Call to action web part]（主）＋[テキスト web part]（補足）",
            "マーケへの業務依頼フォーム", "依頼前にご確認ください：ブランドレビューの手順",
            "ブランドレビュー依頼も含め、マーケへの依頼はすべてこのフォームから。依頼前に手順ガイドで進め方をご確認ください。"
            "　※フォーム（主導線・大きく強調）の下に手順ガイド（副導線・小さくネスト）を配置し、対等な2窓口には見せない構成に変更（T-4対応・前回の誤り訂正）。",
            note="注記: 将来AI-1（セルフチェッカー）／AI-2（アシスタント）の実装により、この依頼導線の構成自体が変わりうる（docs/01 §5-1）。",
            badges=[("T-4", "ok"), ("G-3", "rule")],
        ),
        sec(
            5, "Webサイト流入分析", "[埋め込みWebパーツ web part]",
            "現状のPC表示（埋め込みダッシュボード／iframe相当）をそのまま維持する。",
            note="モバイル専用の分岐表示はWFに描かない（社長指示・R8）。前回WFにあった「SP版のみ表示切替」の記述は本v2から削除。",
            badges=[("R8", "rule")],
        ),
        sec(
            6, "組織体制（配置維持）", "[テキスト web part]（見出しH2）＋組織図画像",
            "見出し「組織体制」をH2で配置しページアンカー生成対象にする。配置は最下部のまま変更しない。目次④からのみ到達性を追加する。",
            badges=[("T-2", "ok"), ("R4", "rule")],
        ),
        changes_box([
            "目次アンカー（Quick Links）を新設し、組織体制含む4項目へのジャンプを明示した（前回不可視だった点の解消）。",
            "モバイル対応（T-1のSP切替表現）はWFから削除し、埋め込みダッシュボードはPC表示を現状維持とした。",
            "組織体制は最下部のまま配置を変更していない（アンカー到達性のみ追加）。",
            "4カテゴリカードに説明文欄を追加した（文言はマーケ部確定待ちのプレースホルダー）。",
            "「業務依頼フォーム」と「ブランドレビューの手順」を対等な2窓口ではなく、フォーム（主）＋依頼前手順ガイド（副）の親子関係として再設計した。",
        ]),
        footnote("ToBe WF v2 — 1. トップ（ホーム） — SharePoint実装可能性(sharepoint-feasibility.md R1〜R9)準拠"),
    ]
    return render("top", 1, 6, "1. トップ（ホーム）", blocks)


# ---------------------------------------------------------------------------
# 2. ブランドライブラリ（1/2: ハブ構成の前半）
# ---------------------------------------------------------------------------
def build_brand_1():
    blocks = [
        hero("ブランドライブラリ", "現状維持"),
        toc([
            {"label": "ブランドの使用について", "target": "注意事項（セクション背景色）"},
            {"label": "ガイドライン一覧", "target": "ガイドラインライブラリ"},
            {"label": "ロゴ掲載申請", "target": "ロゴ掲載申請（2ケース手続き）"},
            {"label": "ロゴ", "target": "ロゴ素材データ", "new": True},
            {"label": "アイコン", "target": "アイコン素材データ", "new": True},
            {"label": "フォント", "target": "フォント"},
            {"label": "画像素材", "target": "画像素材（4サブセクション）"},
            {"label": "各種テンプレート", "target": "各種テンプレート"},
            {"label": "SNSガイドライン", "target": "SNSガイドライン"},
        ], note="前回7項目→9項目（「ロゴ」「アイコン」を新規アンカーとして追加）。パターンはG-1で全ページ共通。"),
        sec(
            3, "ブランドの使用について", "[テキスト web part]（H2見出し1種類に統一）",
            "「ブランドレビュー依頼の前にご確認ください」「注意事項（社外秘）」の2ブロックを<b>セクション背景色（注意喚起トーン）</b>で表現する。"
            "3種類あった黄色マーカー装飾は廃止し、帯見出し1種類に統一する。",
            note="B-6：黄色マーカー→セクション背景色に置換（R7）。大ブロック単位でのセクション折りたたみも可能（R5）。",
            badges=[("R7", "rule"), ("R5", "rule")],
        ),
        cards_block(
            4, "ページ内アンカーカード（旧B-5）", "[クイックリンク web part]",
            [
                ("① バッジ", "→ 「ガイドライン一覧」セクションへのページ内アンカー"),
                ("② 書類スキャン", "→ 「ロゴ掲載申請」セクションへのページ内アンカー"),
                ("③ チャート付き吹き出し", "→ 「各種テンプレート」セクションへのページ内アンカー"),
            ],
            note="社長指摘: 実際はラベル表示ありのページ内アンカーカードだった（スクショが見切れていただけ）。"
            "本ページの構造上、既にラベル付き・アンカー機能付きで運用されている。<b>追加の修正は不要</b>と判定し、"
            "架空の改善（ラベル追加）は作らない。",
            badges=[("判定：修正不要", "ok")],
        ),
        library_block(
            5, "ガイドラインライブラリ（B-1・B-2の核心）",
            "[ドキュメント ライブラリ web part]＋[ハイライトされたコンテンツ web part]",
            ["種別", "言語", "版", "形式", "表示名＝AKKODiSに統一（B-4対応）"],
            [
                ("全社ブランド規定", ["Akkodis ブランドガイドライン V3", "AKKODiS新ビジョンガイドライン"]),
                ("制作物別ガイドライン", ["ビデオガイドライン", "フォトグラフィーガイドライン", "AI生成画像の使用ガイドライン"]),
                ("表記・用語", ["文字表記ガイドライン", "グローバル用語集", "ブランドボイスガイドライン"]),
                ("Akkodis Intelligenceブランド", ["ブランドガイドライン", "Q&A"]),
            ],
            note="目的軸4カテゴリに再構成し、各カテゴリは事前作成ビューで切替表示する。登録前提としてガイドライン10本の重複・旧版の"
            "<b>棚卸し</b>を運用タスクとして実施する（WFのコンポーネントではなく、コンテンツ移行前の運用作業）。",
            altbox="ハブ&スポーク構成（主案）: 本セクションは情報量が多いため、ハブページ（本ページ）には要約カード＋詳細への導線のみを置き、"
            "詳細は<b>同一サイト内の別ページ（スポーク）</b>として切り出す（2枚目WF「brand_2」相当）。"
            "副案: ページ分割をせず、<b>ハブページ内でのセクション折りたたみ格納（大ブロック単位・既定は折りたたみ状態）</b>とする代替構成。"
            "どちらを採るかは要テナント確認／要社長判断。",
            badges=[("B-1", "ok"), ("B-2", "ok"), ("R3", "rule"), ("R5", "rule"), ("R6", "rule")],
        ),
        changes_box([
            "「ロゴ」セクションと「アイコン」セクションを新設し、既存の「ロゴ掲載申請」（手続き）・冒頭アンカーカード（旧B-5）とは別物であることをラベルで明記した（brand_2で詳述）。",
            "B-1（情報過多）は目次追加だけでなく、ハブ&スポーク構成（ページ分割 or 折りたたみ格納）＋棚卸し＋要約カードで情報量そのものを削減する設計に変更した。",
            "B-2（ガイドライン10本）は目的軸4カテゴリのドキュメントライブラリ＋ビュー切替に再構成した。",
            "B-5は実際には修正不要と判定し、架空の改善を作らなかった（前回は「ラベル追加」としていたが撤回）。",
            "B-6（黄色マーカー）はセクション背景色による注意喚起表現に置き換え、色は固定パレット内で全ページ統一する方針にした。",
        ]),
        footnote("ToBe WF v2 — 2. ブランドライブラリ（1/2・ハブ前半） — 続きは brand_2 参照"),
    ]
    return render("brand_1", 2, 6, "2. ブランドライブラリ（1/2）", blocks)


# ---------------------------------------------------------------------------
# 2. ブランドライブラリ（2/2: ロゴ掲載申請〜SNSガイドライン）
# ---------------------------------------------------------------------------
def build_brand_2():
    blocks = [
        sec(
            "（続き）", "ブランドライブラリ（2/2）— ハブ後半／スポーク相当", "—",
            "brand_1（目次・ブランドの使用について・アンカーカード・ガイドラインライブラリ）からの続き。"
            "「ロゴ掲載申請」以降のセクションを掲載する。",
        ),
        cols2_block(
            6, "ロゴ掲載申請（手続き）", "[テキスト web part]（番号付きリスト）＋[ドキュメント ライブラリ web part]（同意書テンプレ）",
            ("①先方企業のロゴを当社サイトに掲載する場合",
             ["Marketing部依頼フォーム提出", "Teams連絡", "Salesforceサービスリクエストから押印申請", "回収", "送付", "HP掲載"],
             "同意書テンプレ（先方ロゴ掲載用）"),
            ("②当社社名・ロゴを他社サイト等に掲載する場合",
             ['<span class="tbd">（マーケ部確定待ち・例）</span>手順は現物確認の上マーケ部が確定',
              "構造要件のみ規定：①と同様にフォーム起点の番号付き手順として構造化する（具体手順は創作しない）"],
             "同意書テンプレ（当社ロゴ掲載用）"),
            note="2カラムはSharePointの<b>セクション内2列レイアウト</b>（標準機能）で実現。タブUIではなく静止した2カラム表示とする。"
            "B-3：長文6ステップ→2ケースに整理し番号付き手順で構造化（前回「不明」指摘への回答）。",
            badges=[("B-3", "ok"), ("R5", "rule")],
        ),
        library_block(
            7, "ロゴ（新設）", "[ドキュメント ライブラリ web part]",
            ["用途", "バリエーション", "カラー", "形式：AI／EPS／PNG／SVG等"],
            [("ロゴ素材データ", ["AKKODiSロゴ（横組み）", "AKKODiSロゴ（縦組み）", "モノクロ版", "Akkodis Intelligence版"])],
            note='ロゴ素材データはこちら。他社サイトへの掲載申請は「ロゴ掲載申請」セクションへ（別物である旨を帯見出し直下に明記）。'
            "前回WFで欠落していたセクション。",
            badges=[("新設", "newf"), ("R3", "rule")],
            new=True,
        ),
        sec(
            8, "フォント", "[テキスト web part]",
            "現状のテキスト構成を維持する。目次から到達可能にするのみで内容変更はなし。",
        ),
        cards_block(
            9, "画像素材（整理・構造化）", "[テキスト web part]＋[クイックリンク web part]（4サブセクション）",
            [
                ("マーケティング部承認済み画像について", "1〜2行要約。詳細は本カードのアンカー先段落へ"),
                ("画像の購入について", "1〜2行要約。詳細は本カードのアンカー先段落へ"),
                ("AI生成画像の使用について", "1〜2行要約。詳細は本カードのアンカー先段落へ"),
                ("使用禁止の画像について", "1〜2行要約。詳細は本カードのアンカー先段落へ"),
            ],
            note="現状の長文4見出しを、サブ見出し4枚のカード＋各カード1〜2行要約＋詳細リンクの構造に変換（B-1関連：長文→カード化で走査性改善）。",
            badges=[("R4", "rule")],
        ),
        library_block(
            10, "アイコン（新設）", "[ドキュメント ライブラリ web part]",
            ["用途別"],
            [("アイコン素材データ", ["UI用アイコン", "資料用アイコン", "SNS用アイコン"])],
            note="本セクションはアイコン素材データ。ページ内ナビ用のアイコンカードは冒頭の④参照（旧B-5のアンカーカードとの混同を解消）。"
            "前回WFで欠落していたセクション。",
            badges=[("新設", "newf"), ("R3", "rule")],
            new=True,
        ),
        sec(
            11, "各種テンプレート", "[ハイライトされたコンテンツ web part]",
            "コンテンツタイプ別（PowerPoint／Word／メール署名／Forms／バーチャル背景）。絞り込み文言はR2の定型表現に統一：<br/>"
            "「<b>[ハイライトされたコンテンツ web part] で、事前作成ビュー／Highlighted contentのType指定／管理プロパティ指定により絞り込み"
            "（テナントの検索スキーマ・クロールドプロパティのマッピングに依存・要テナント確認）</b>」。"
            "「ワンクリックで動的に多条件フィルタできる」という誤った表現は使わない。",
            badges=[("R2", "rule")],
        ),
        cards_block(
            12, "SNSガイドライン", "[クイックリンク web part]",
            [("SNSガイドライン", "現状維持")],
        ),
        future_block(
            "AI検索の高度ファセット／Copilot要約等",
            "L1範囲外のためMVPワイヤー本体には含めない。ファセット・リファイナー等の高度な絞り込みはL3（PnP Modern Search等・カスタムコード許可が前提）。",
        ),
        changes_box([
            "「ロゴ」セクション（新設）と「アイコン」セクション（新設）を追加し、それぞれ「ロゴ掲載申請」（手続き）・冒頭アンカーカードとは別物と明記した。",
            "B-3（ロゴ掲載6ステップ）を「先方ロゴ掲載」「当社ロゴ掲載」の2ケースに整理し、番号付き手順・2カラムで構造化した。②の具体手順は創作せず「担当確認待ち・暫定」の構造のみ規定。",
            "「画像素材」の長文4見出しをカード＋要約＋詳細リンクの構造に整理した。",
            "「各種テンプレート」のコンテンツタイプ絞り込みは、R2に基づく正確な文言（事前ビュー/Type/管理プロパティ・テナント検索設定依存）に修正した。",
            "B-4（表記ゆれ）はメタデータ列・表示名をAKKODiS表記に統一する設計要件を明記した。",
        ]),
        footnote("ToBe WF v2 — 2. ブランドライブラリ（2/2） — 前半は brand_1 参照"),
    ]
    return render("brand_2", 3, 6, "2. ブランドライブラリ（2/2）", blocks)


# ---------------------------------------------------------------------------
# 3. クライアント向け
# ---------------------------------------------------------------------------
def build_client():
    blocks = [
        hero("クライアント向け", "現状維持"),
        toc([
            {"label": "AKKODiS企業向けWebページ", "target": "セミナー／活用事例／Insights／サービス"},
            {"label": "公開資料ライブラリ", "target": "Lists（カテゴリ・言語・更新日）"},
            {"label": "メルマガ配信スケジュール", "target": "配信スケジュール（社外秘）"},
            {"label": "資料を探す", "target": "", "removed": True},
        ]),
        rationale(
            "「資料を探す」アンカー削除の判断根拠（3-3）",
            "前回WFは「資料を探す（新規）」という4つ目のアンカーを追加していたが、遷移先が明記されておらず社長から指摘を受けた。"
            "精査の結果、意図は①本文検索、または②公開資料ライブラリへのショートカットのいずれかと推測されるが、"
            "①はR1により実装不可、②は既存の目次項目「公開資料ライブラリ」と完全重複する。したがって<b>曖昧な項目として削除</b>し、"
            "代わりに「公開資料ライブラリ」セクション自体にLists標準のビュー/列フィルター機能を持たせることで同じニーズに応える設計とした。",
        ),
        cards_block(
            3, "AKKODiS企業向けWebページ", "[クイックリンク web part]",
            [("セミナー", "現状維持"), ("活用事例", "現状維持"), ("AKKODiS Insights", "現状維持"), ("サービスページ", "現状維持")],
        ),
        library_block(
            4, "公開資料ライブラリ（旧・ホームページ公開資料）", "[Lists web part]（カスタムリスト・ハイパーリンク列）",
            ["タイトル＝ハイパーリンク列", "カテゴリ＝選択肢列", '<span class="tbd">事例資料＝新規追加</span>', "言語＝選択肢列", "更新日＝任意列（付与しない資料があってよい）"],
            [
                ("共通", ["会社案内（日本語版）", "会社案内（英語版）"]),
                ("ソリューション", ["紹介資料 約11本"]),
                ("アカデミー", ["紹介資料 約8本"]),
                ("事例資料（新規）", ['<span class="tbd">（マーケ部確定待ち・例）</span>掲載予定資料は今後選定']),
                ("お役立ち資料", ["30本超（日本語版／英語版は言語列で1エントリに集約）"]),
            ],
            note="C-1：資料の実体は外部AKKODiSサイトへのリンクであり、ファイルではないため<b>ドキュメントライブラリ→Listsに設計変更</b>（R6準拠）。"
            "C-2：更新日は必須列にせず、任意列として「あえて付与しない」現状の方針を尊重する。",
            badges=[("C-1", "ok"), ("C-2", "ok"), ("R6", "rule")],
        ),
        sec(
            5, "「ページ内資料の探し方」案内（新設・検索代替）", "[テキスト web part]",
            "①画面上部の検索でキーワード検索　②このリストの列見出し（カテゴリ／言語）をクリックして絞り込み（sort/filter/group）",
            note="ブラウザ内検索（Ctrl+F）とは異なる旨を注記する。「このページ内資料検索」（検索ボックス新設）はR1により実装不可のため削除し、"
            "本案内に置換した。",
            badges=[("R1", "rule"), ("R3", "rule")],
        ),
        cards_block(
            6, "メルマガ配信スケジュール", "[クイックリンク web part]",
            [("メルマガ配信スケジュール", "現状維持（社外秘表記も維持）")],
        ),
        changes_box([
            "「公開資料ライブラリ」のカテゴリ列に「事例資料」を追加した。",
            "「資料を探す」アンカーは遷移先が特定できないため削除し、既存の目次3項目に整理した（重複回避・判断根拠は上部参照）。",
            "「このページ内資料検索」はR1により実装不可のため削除し、ヘッダー既定検索＋Listsのビュー/列フィルターの組み合わせに置換した。",
            "C-1は「メタデータ列付きドキュメントライブラリ」という前回の誤りを修正し、Lists（ハイパーリンク列＋ビュー）に設計変更した。",
            "C-2は更新日を必須列にせず、任意列として「あえて付与しない」現状の方針を尊重する設計にした。",
        ]),
        footnote("ToBe WF v2 — 3. クライアント向け"),
    ]
    return render("client", 4, 6, "3. クライアント向け", blocks)


# ---------------------------------------------------------------------------
# 4. キャンディデート向け
# ---------------------------------------------------------------------------
def build_candidate():
    blocks = [
        hero("キャンディデート向けページ", "現状維持"),
        toc([
            {"label": "AKKODiS公式HP", "target": "AKKODiS People／リクルーティングムービー／業務依頼フォーム"},
            {"label": "マスター資料", "target": "会社説明資料（直リンク・暫定）"},
            {"label": "Social Media", "target": "X／Facebook／LinkedIn／YouTube"},
        ]),
        cards_block(
            2, "AKKODiS公式HP", "[クイックリンク web part]",
            [
                ("AKKODiS People", "現状維持"),
                ('業務依頼フォームはこちら<span class="new" style="margin-left:6px;">新規追加</span>',
                 '<span class="tbd">（マーケ部確定待ち・例）</span>「AKKODiS People」項目の直下（または並び）に新規追加'),
                ("リクルーティングムービー", "現状維持"),
            ],
            note="「AKKODiS People」に業務依頼フォームへのリンクを新規追加した（社長指摘の反映）。",
            badges=[("新規追加", "newf")],
        ),
        library_block(
            3, "マスター資料（CA-1・暫定）", "[ドキュメント ライブラリ web part]＋[クイックリンク web part]（直リンク）",
            ["資料オーナー／問い合わせ先（現状の3チーム連絡先をメタデータとして保持）"],
            [("会社説明資料（直リンク）", ["中途採用向け", "新卒採用向け", "障がい者採用向け"])],
            note='<span class="tbd">担当確認待ち・暫定案（社長確認中）</span> — CA-1（マスター資料の直リンク化案）は社長が担当者確認中のため、'
            "確定案としては描いていない。承認が下り次第、確定ラベルに差し替える。",
            badges=[("CA-1・保留", "ok")],
        ),
        cards_block(
            4, "Social Media", "[クイックリンク web part]",
            [("X", "現状維持"), ("Facebook", "現状維持"), ("LinkedIn", "現状維持"), ("YouTube", "現状維持")],
        ),
        changes_box([
            "「AKKODiS公式HP」セクションの「AKKODiS People」に業務依頼フォームへのリンクを追加した。",
            "CA-1（マスター資料の直リンク化案）は社長が担当者確認中のため、WF上で「担当確認待ち・暫定案」と明記し、確定案としては描いていない。",
            "CA-2（3セクション構成の簡潔さ）は変更していない。",
        ]),
        footnote("ToBe WF v2 — 4. キャンディデート向け"),
    ]
    return render("candidate", 5, 6, "4. キャンディデート向け", blocks)


# ---------------------------------------------------------------------------
# 5. グローバルブランドパートナーシップ
# ---------------------------------------------------------------------------
def build_partner():
    blocks = [
        hero("グローバルブランドパートナーシップ", "ヒーロー＋リード文（現状維持）"),
        cards_block(
            2, "3パートナーのサムネイルカード（役割A）", "[クイックリンク web part]（タイル/グリッドレイアウト）",
            [
                ("Mercedes-AMG PETRONAS Formula One Team", "＝各パートナー紹介へのご案内カード"),
                ("Akkodis ASP Team", "＝各パートナー紹介へのご案内カード"),
                ("Stade Toulousain", "＝各パートナー紹介へのご案内カード"),
            ],
            note="各カードの役割は「そのパートナー紹介への誘導（クリックで下部の該当セクションへ）」。大きめのタイル/グリッド（画像＋パートナー名）。",
        ),
        toc([
            {"label": "Mercedes-AMG PETRONAS Formula One Team", "target": "該当パートナーセクション"},
            {"label": "Akkodis ASP Team", "target": "該当パートナーセクション"},
            {"label": "Stade Toulousain", "target": "該当パートナーセクション"},
        ], note="＝ページ内ジャンプ用の目次（役割B）。役割は「同一ページ内の各パートナーセクションへのジャンプ」専用。コンパクトなテキストリンク帯で、役割Aのカードとは視覚的に区別する。"),
        cmp_table(
            "サムネイル（役割A）と目次（役割B）の違い（5-3・前回「重複では」との指摘への回答）",
            [
                ("目的", "各パートナーの「紹介」への誘導（一覧性・ビジュアル訴求）", "同一ページ内の該当セクションへの<b>ジャンプのみ</b>"),
                ("見た目", "大きめのタイル/グリッド（画像＋パートナー名）", "コンパクトなテキストリンクの並び（帯状）"),
                ("配置", "ヒーロー直下（#2）", "サムネイルの直後、各パートナーセクションの直前（#3）"),
                ("SharePointパーツ", "クイックリンク web part（タイル/グリッド）", "クイックリンク web part（リスト/コンパクト）"),
            ],
        ),
        sec(
            4, "各パートナーセクション（3ブロック）", "[テキスト web part]（H2見出し）＋既存構成",
            "概要／共通価値／パートナーシップビデオ（現状維持）。各ブロック先頭の見出しが目次アンカー先になる。",
            badges=[("R4", "rule")],
        ),
        sec(
            5, "Akkodis ASP Teamの画像枠（P-1・修正済み）", "[画像 web part]",
            "他2チームと同じ構造に統一済み（維持）。",
        ),
        changes_box([
            "サムネイルカード（誘導用・ビジュアル一覧）とページ内目次（ジャンプ専用・テキストリンク）の役割の違いをレイアウト・注記・比較表で明確化した。",
            "P-1（ASP Teamの画像枠是正）・P-2（ページ内目次新設）は前回修正済みのため、本v2でも維持している。",
        ]),
        footnote("ToBe WF v2 — 5. グローバルブランドパートナーシップ"),
    ]
    return render("partner", 6, 6, "5. グローバルブランドパートナーシップ", blocks)


# ---------------------------------------------------------------------------
# PDF結合
# ---------------------------------------------------------------------------
def build_pdf(png_paths, out_name="ToBe_WF_v2_承認用.pdf"):
    imgs = [Image.open(p).convert("RGB") for p in png_paths]
    out_path = os.path.join(OUT, out_name)
    imgs[0].save(out_path, save_all=True, append_images=imgs[1:])
    print("saved", out_path)
    return out_path


if __name__ == "__main__":
    paths = []
    paths.append(build_top())
    paths.append(build_brand_1())
    paths.append(build_brand_2())
    paths.append(build_client())
    paths.append(build_candidate())
    paths.append(build_partner())
    build_pdf(paths)
    print("all done")

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ToBe ワイヤーフレーム v3 生成スクリプト（判読性最優先・大判レンダリング）
- v2（build_tobe_wf_v2.py）を土台に、社長2回目差戻し＋一次情報再検証
  （sharepoint-verify-v3.md・ai12-implementation.md）を反映した改訂版。
- 主な変更点（詳細は wf-v3-directives.md）:
  1) 各ページ＝1枚（ブランドの1/2・2/2を統合）。計5枚。
  2) カラムは最大3列（4列があれば行を分ける＝CSS gridで自動強制）。
  3) パーツ正式名称に統一（強調表示されたコンテンツ web part／リスト web part／
     テキスト web part／このセクションを折りたたみ可能にする）。
  4) AI-1（ブランドセルフチェッカー）／AI-2（サイト共通エージェント）を
     実装要素として実線枠で組込（ライセンス/権限要確認の小注記つき）。
  5) ブランド：Section4アンカーカード削除／B-1折りたたみ格納／B-2新規ライブラリ
     不作成（テキストwebパーツでリンク）／「ブランドの使用について」詳細化／
     SNSガイドラインはテキストwebパーツ／ロゴ・アイコンは既存素材リンク。
  6) クライアント：公開資料ライブラリをリストwebパーツ（行＋列）で3列以内に。
     Section5は独立セクションをやめ枠内1行注記に統合。
  7) キャンディデート：AKKODiS People業務依頼リンクは文脈内リンクに。
     CA-1は現行文言をそのままテキストwebパーツに掲載（暫定ラベル撤去）。
  8) パートナー：サムネイルカード削除・目次（ページ内リンク）のみ残す。
"""
import os
from playwright.sync_api import sync_playwright
from PIL import Image

OUT = "/workspace/akkodis-marketing-portal/deliverables/1784786903168-17720/wf_v3"
os.makedirs(OUT, exist_ok=True)

CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

WIDTH = 1160  # 本文幅（v2の1120pxからやや拡張。3列グリッド化で余白を確保）
TOTAL_PAGES = 5

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

/* ページ名見出し */
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
  border-bottom:1px solid #2c333c;
}
.searchband .tag {
  background:#6c7480; border:1px solid #939ba6; color:#fff;
  padding:3px 10px; border-radius:3px; font-size:11.5px; font-weight:bold;
}
.searchband .desc { color:#cfd4da; font-size:12.5px; }

/* AI-2 帯（サイト共通エージェント・全ページ共通・実装要素＝実線） */
.ai2band {
  background:#eaf3ec; border-bottom:3px solid #2f6b3f; padding:10px 26px;
  display:flex; align-items:center; gap:12px; font-size:12.5px; color:#1c4a28;
  flex-wrap:wrap;
}
.ai2band .tag {
  background:#2f6b3f; color:#fff; padding:3px 10px; border-radius:3px;
  font-size:11px; font-weight:bold; white-space:nowrap;
}
.ai2band .body { color:#1c4a28; }
.ai2band .lic { font-size:10.5px; color:#5a7d63; margin-left:auto; }

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
.sec .badges { display:flex; gap:6px; flex-wrap:wrap; justify-content:flex-end; max-width:280px; }
.badge { font-size:11px; padding:2px 8px; border-radius:3px; border:1px solid #999; color:#444; background:#eee; white-space:nowrap;}
.badge.rule { background:#e7edf9; border-color:#7d94c9; color:#28407e; }
.badge.newf { background:#e3f3e6; border-color:#4f9463; color:#1d5c30; }
.badge.ok { background:#e6f0fb; border-color:#5c85b8; color:#204a78; }
.badge.v3 { background:#fbe6f4; border-color:#c23b82; color:#7a1f52; font-weight:bold; }
.badge.removed { background:#f3e6e6; border-color:#b85c5c; color:#7a2020; }
.desc { font-size:14.5px; line-height:1.85; color:#2b2b2b; }
.note {
  margin-top:10px; font-size:13px; color:#555; line-height:1.75;
  border-left:3px solid #9aa; padding-left:10px;
}
.tbd {
  color:#8a4b00; background:#fff3de; border:1px solid #d69a3f;
  padding:1px 7px; border-radius:3px; font-size:13px; font-weight:bold;
}

/* AI-1 実装ボックス（実線・実装要素） */
.ai1box {
  margin-top:12px; border:2px solid #1c5f8a; background:#eaf3fa; padding:14px 18px;
}
.ai1box .tag { display:inline-block; font-size:11px; font-weight:bold; color:#fff; background:#1c5f8a; padding:2px 8px; border-radius:4px; margin-bottom:8px;}
.ai1box .t { font-size:14.5px; font-weight:bold; color:#123b54; margin-bottom:6px; }
.ai1box .d { font-size:13px; color:#1c455e; line-height:1.75; }
.ai1box .lic { margin-top:9px; font-size:10.5px; color:#6b5320; background:#fff7e8; border:1px dashed #c9a15a; padding:5px 9px; border-radius:3px; }

/* 主導線/副導線ボタン風 */
.cta-primary {
  display:inline-block; background:#2b3a55; color:#fff; font-size:16px; font-weight:bold;
  padding:14px 26px; border-radius:5px; border:2px solid #1c2733;
}
.cta-secondary {
  display:inline-block; margin-top:4px; font-size:13.5px; color:#2b3a55;
  border-bottom:1.5px solid #7f8aa0; padding-bottom:2px;
}
.cta-relation { font-size:13px; color:#555; margin:12px 4px 4px; line-height:1.8; }

/* T-4: ①AI-1 → ②手順ガイド → ③業務依頼フォーム の3ステップ導線（3列以内） */
.flow3 { display:flex; align-items:stretch; gap:10px; margin-top:10px; }
.flow3 .step { flex:1; display:flex; flex-direction:column; }
.flow3 .stepnum {
  font-size:12px; font-weight:bold; color:#fff; background:#555; width:22px; height:22px;
  border-radius:50%; display:flex; align-items:center; justify-content:center; margin-bottom:8px;
}
.flow3 .arrow { display:flex; align-items:center; justify-content:center; font-size:22px; color:#8a8f96; flex:0 0 24px; }
.flow3 .guidebox { border:2px solid #8a8f96; background:#fff; padding:14px 16px; flex:1; }
.flow3 .guidebox .t { font-size:13.5px; font-weight:bold; color:#1e2126; margin-bottom:6px; }
.flow3 .guidebox .d { font-size:12px; color:#555; line-height:1.7; }
.flow3 .formbox { border:2px solid #2b3a55; background:#eef1f6; padding:14px 16px; flex:1; display:flex; flex-direction:column; justify-content:center; align-items:flex-start; }

/* カード群（V1: 3列以内・グリッドで自動的に4件目以降を次の行へ折り返す） */
.cards { display:grid; grid-template-columns: repeat(3, 1fr); gap:14px; margin-top:6px; }
.card {
  border:2px solid #8a8f96; background:#fff; padding:14px 14px;
}
.card .icon { width:34px; height:34px; background:#c9cbce; border-radius:6px; margin-bottom:8px; }
.card .t { font-size:14.5px; font-weight:bold; color:#1e2126; margin-bottom:6px; }
.card .d { font-size:12.5px; color:#555; line-height:1.6; }
.card .inline-link { display:inline-block; margin-top:6px; font-size:12px; color:#1c5f8a; border-bottom:1px solid #1c5f8a; padding-bottom:1px; }

/* 2カラム */
.cols2 { display:flex; gap:18px; margin-top:8px; }
.col { flex:1; border:1.5px solid #9a9fa6; background:#fff; padding:14px 16px; }
.col .t { font-size:14.5px; font-weight:bold; margin-bottom:8px; color:#1e2126; }
.col ol { margin-left:18px; font-size:13px; line-height:2; color:#2b2b2b; }
.col .link { margin-top:10px; font-size:12px; color:#2b3a55; border-top:1px dashed #aaa; padding-top:8px; }

/* サブブロック（ブランドの使用について・詳細化） */
.subblocks { display:flex; flex-direction:column; gap:12px; margin-top:10px; }
.subblock { border-left:6px solid #999; padding:12px 18px; background:#fafafa; }
.subblock.tone-a { background:#fff4e0; border-color:#c58a2e; }
.subblock.tone-b { background:#fdeceb; border-color:#c0392b; }
.subblock .st { font-size:14px; font-weight:bold; margin-bottom:6px; color:#222; }
.subblock .sd { font-size:12.8px; color:#3a3a3a; line-height:1.8; }

/* 折りたたみ（V4: 「このセクションを折りたたみ可能にする」） */
.collapsible { margin-top:12px; border:2px solid #55606b; background:#fff; }
.collapse-head {
  background:#e8eaed; color:#2a2f36; font-size:13px; font-weight:bold;
  padding:10px 16px; border-bottom:2px solid #55606b; display:flex; justify-content:space-between; align-items:center;
}
.collapse-head .st { font-size:11px; font-weight:normal; color:#666; }
.collapse-body { padding:14px 18px; }

/* リンク一覧（テキストweb部品：ロゴ/アイコン/SNS/ガイドライン） */
.linkgroups { display:grid; grid-template-columns: repeat(2, 1fr); gap:14px; margin-top:8px; }
.linkgroups.g1 { grid-template-columns: 1fr; }
.linkgroup { border:1.5px solid #9a9fa6; background:#fbfbfc; padding:12px 16px; }
.linkgroup .n { font-size:13.5px; font-weight:bold; margin-bottom:8px; color:#1e2126; }
.linkgroup ul { list-style:none; }
.linkgroup li { font-size:12.5px; color:#1c5f8a; padding:4px 0 4px 16px; position:relative; line-height:1.65; text-decoration:underline; }
.linkgroup li:before { content:"🔗"; position:absolute; left:0; font-size:10px; top:5px; }

/* リスト web part（クライアント公開資料ライブラリ・行＋列。3列以内=単一ワイドテーブル） */
.listpart { margin-top:10px; border:1.5px solid #8a8f96; background:#fff; }
.listpart .filterbar { display:flex; gap:8px; flex-wrap:wrap; padding:10px 14px; background:#eef1f6; border-bottom:1.5px solid #8a8f96; align-items:center; }
.listpart .filterbar .lbl { font-size:11.5px; color:#444; font-weight:bold; margin-right:4px; }
.listpart .chip { font-size:11.5px; background:#fff; border:1.5px solid #3b5bb0; color:#1c3570; padding:3px 11px; border-radius:14px; }
.listpart table { width:100%; border-collapse:collapse; }
.listpart th, .listpart td { border:1px solid #d3d5d8; padding:8px 12px; font-size:12.5px; text-align:left; vertical-align:top; line-height:1.55; }
.listpart th { background:#e9ebee; font-size:12px; }
.listpart td.linkcol { color:#1c5f8a; text-decoration:underline; }
.listpart td.catcell { white-space:nowrap; }
.listpart .footnote1 { padding:10px 14px; font-size:12px; color:#444; background:#f7f7f8; border-top:1.5px solid #d3d5d8; }

/* ライブラリメタ列 */
.lib-meta { display:flex; gap:8px; flex-wrap:wrap; margin:10px 0 0; }
.lib-meta span {
  font-size:11.5px; background:#e2e4e8; border:1px solid #9a9fa6; color:#333;
  padding:3px 9px; border-radius:3px;
}

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
# 共通ビルダー関数群
# ---------------------------------------------------------------------------
def titlebar(name, idx, total, sub=""):
    return f'''<div class="titlebar">
      <div class="name">ToBe ワイヤーフレーム v3　｜　{name}　（{idx}/{total}）</div>
      <div class="sub">{sub or "AKKODiSマーケティングポータル・承認用WF（社長2回目差戻し反映）"}</div>
    </div>'''


def searchband():
    return '''<div class="searchband">
      <span class="tag">システム標準・編集不可</span>
      <span>サイトヘッダー：Microsoft Search（既定表示・位置固定・全ページ共通）</span>
      <span class="desc">※本文コンポーネントではない。ページ本文には検索ボックスを配置しない（R1）</span>
    </div>'''


def ai2band():
    return '''<div class="ai2band">
      <span class="tag">実装要素・全ページ共通</span>
      <span class="body">サイト共通エージェント：<b>AI-2</b>（SharePoint agent・ヘッダーの「エージェント」アイコンから起動・利用者本人のアクセス権範囲でのみ回答・出典/引用リンク付き）</span>
      <span class="lic">※要 Microsoft 365 Copilotライセンス or Pay-as-you-go課金（要テナント管理者確認）</span>
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


def ai1box(context_label="ブランドレビュー依頼／ロゴ掲載申請"):
    return f'''<div class="ai1box">
      <span class="tag">実装要素：AI-1（{context_label}の文脈に組込）</span>
      <div class="t">[Embed web part] AI-1（ブランドセルフチェッカー）</div>
      <div class="d">Copilot Studioエージェントを Embed web part でページに埋め込み。PPT/画像をアップロードするとブランドガイドラインと
      照合した<b>一次スクリーニング</b>（気づきの提示）を返す。色/ロゴの厳密なピクセル判定は保証しない。<b>最終承認は人が行う</b>（自動承認はしない）。</div>
      <div class="lic">※Copilot Studioライセンス／PPTXアップロードは実験的機能につき要Microsoft申請／カスタムWebサイトチャネル公開は認証なし設定が前提
      → いずれも要テナント管理者確認</div>
    </div>'''


def collapsible(head_note, inner_html, default_state="既定：折りたたみ"):
    return f'''<div class="collapsible">
      <div class="collapse-head"><span>▼ このセクションを折りたたみ可能にする（{default_state}）</span><span class="st">{head_note}</span></div>
      <div class="collapse-body">{inner_html}</div>
    </div>'''


def linkgroups_block(groups, single_col=False):
    cls = "linkgroups g1" if single_col else "linkgroups"
    gs = "".join(
        f'<div class="linkgroup"><div class="n">{g[0]}</div><ul>{"".join(f"<li>{x}</li>" for x in g[1])}</ul></div>'
        for g in groups
    )
    return f'<div class="{cls}">{gs}</div>'


def future_block(title, desc):
    return f'''<div class="future">
      <span class="tag">将来拡張（L2/L3・本体には含めない）</span>
      <div class="t">{title}</div>
      <div class="d">{desc}</div>
    </div>'''


def changes_box(lines, head="このページの変更点（v3・社長2回目差戻し反映）"):
    lis = "".join(f"<li>{l}</li>" for l in lines)
    return f'''<div class="changes">
      <div class="h"><span class="pin">v3変更点</span>{head}</div>
      <ul>{lis}</ul>
    </div>'''


def footnote(text):
    return f'<div class="footnote">{text}</div>'


def render(name, idx, total, page_label, blocks, sub=""):
    css = CSS.replace("__WIDTH__", str(WIDTH))
    body = (titlebar(page_label, idx, total, sub) + searchband() + ai2band()
            + '<div class="page">' + "".join(blocks) + '</div>')
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
            {"label": "②マーケへの業務依頼", "target": "AI-1セルフチェック→手順ガイド→業務依頼フォーム"},
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
            note="各カード直下に1行の説明文欄（T-3対応・v2から変更なし）。カラムは3列グリッドのため4枚目は次の行に自動で折り返す（V1）。",
            badges=[("T-3", "ok")],
        ),
        (lambda: f'''<div class="sec new">
          <div class="sechead">
            <div>
              <div class="num">SECTION 4</div>
              <div><span class="part-tag">①[Embed web part]　②[テキスト web part]　③[Call to action web part]</span></div>
              <div class="title">マーケへの業務依頼（v3再設計・AI-1組込）</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span><span class="badge ok">T-4</span><span class="badge rule">G-3</span></div>
          </div>
          <div class="desc">依頼前にAI-1で自己チェック→手順ガイドを確認→業務依頼フォーム送信、という<b>3ステップの導線</b>に再設計する
          （ブランドレビュー依頼もこのフォームの依頼種別の1つ。対等な2窓口ではなくフォーム主導線＋手順ガイド副導線の親子関係はv2から維持）。</div>
          <div class="flow3">
            <div class="step">
              <div class="stepnum">①</div>
              {ai1box("マーケへの業務依頼（依頼前セルフチェック）")}
            </div>
            <div class="arrow">→</div>
            <div class="step">
              <div class="stepnum">②</div>
              <div class="guidebox">
                <div class="t">[テキスト web part] 手順ガイド</div>
                <div class="d">「依頼前にご確認ください：ブランドレビューの手順」。フォームより小さい副導線・確認用リンクとして配置（v2から維持）。</div>
              </div>
            </div>
            <div class="arrow">→</div>
            <div class="step">
              <div class="stepnum">③</div>
              <div class="formbox">
                <div class="cta-primary">[Call to action web part]<br/>マーケへの業務依頼フォーム</div>
              </div>
            </div>
          </div>
          <div class="cta-relation">導線の関係説明: 「①AI-1で依頼前に自己チェック（気づきの提示）→②手順ガイドで進め方を確認→③フォームから正式に依頼」。
          対等な3枚のカードではなく、左から右へ進む<b>一本の依頼フロー</b>として描く。</div>
          <div class="note">3ステップは3列以内（V1準拠）。AI-1／手順ガイドはあくまで依頼前の確認・自己チェック導線であり、最終的な依頼窓口はフォーム1つに統一されている（G-3）。</div>
        </div>''')(),
        sec(
            5, "Webサイト流入分析", "[埋め込みWebパーツ web part]",
            "現状のPC表示（埋め込みダッシュボード／iframe相当）をそのまま維持する。",
            note="モバイル専用の分岐表示はWFに描かない（社長指示・R8）。v2から変更なし。",
            badges=[("R8", "rule")],
        ),
        sec(
            6, "組織体制（配置維持）", "[テキスト web part]（見出しH2）＋組織図画像",
            "見出し「組織体制」をH2で配置しページアンカー生成対象にする。配置は最下部のまま変更しない。目次④からのみ到達性を追加する。",
            badges=[("T-2", "ok"), ("R4", "rule")],
        ),
        changes_box([
            "【v3】SECTION4「マーケへの業務依頼」に AI-1（ブランドセルフチェッカー）を実装要素として組込み、"
            "「①AI-1自己チェック→②手順ガイド→③業務依頼フォーム」の3ステップ導線に再設計した（点線の将来拡張ではなく実線の実装要素。ライセンス/権限の要確認事項は小注記で明記）。",
            "【v3】全ページ共通でヘッダー直下にAI-2（サイト共通エージェント／SharePoint agent）の実装注記帯を追加した。",
            "（v2からの継続）目次アンカー（Quick Links）を新設し、組織体制含む4項目へのジャンプを明示。モバイル対応（T-1）はWFから削除。4カテゴリカードに説明文欄を追加。",
        ]),
        footnote("ToBe WF v3 — 1. トップ（ホーム） — SharePoint実装可能性 v3再検証（sharepoint-verify-v3.md）／AI実装調査（ai12-implementation.md）準拠"),
    ]
    return render("top", 1, TOTAL_PAGES, "1. トップ（ホーム）", blocks)


# ---------------------------------------------------------------------------
# 2. ブランドライブラリ（v3: 1枚に統合）
# ---------------------------------------------------------------------------
def build_brand():
    blocks = [
        hero("ブランドライブラリ", "現状維持"),
        toc([
            {"label": "ブランドの使用について", "target": "サブブロックA/B（注意喚起・AI-1チェック導線）"},
            {"label": "ガイドライン一覧", "target": "テキストwebパーツ・4カテゴリ見出し＋リンク箇条書き（折りたたみ格納）"},
            {"label": "ロゴ掲載申請", "target": "2ケース手続き（AI-1組込）"},
            {"label": "ロゴ", "target": "既存ロゴ素材へのリンク"},
            {"label": "アイコン", "target": "既存アイコン素材へのリンク"},
            {"label": "フォント", "target": "フォント"},
            {"label": "画像素材", "target": "画像素材（4サブセクション）"},
            {"label": "各種テンプレート", "target": "強調表示されたコンテンツ web part"},
            {"label": "SNSガイドライン", "target": "テキストwebパーツ・リンク箇条書き"},
        ], note="9アンカー（v2から項目数は変更なし）。<b>旧「Section4 ページ内アンカーカード」は上部目次と機能重複のため削除した（v3）</b>。"),
        (lambda: f'''<div class="sec new">
          <div class="sechead">
            <div>
              <div class="num">SECTION 3</div>
              <div><span class="part-tag">[テキスト web part]（H2見出し1種類に統一）</span></div>
              <div class="title">ブランドの使用について（v3・サブブロックに分けて詳細化）</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span><span class="badge rule">R7</span></div>
          </div>
          <div class="desc">1つの長文セクションではなく、<b>役割の異なる2つのサブブロック</b>に分けて描く（社長指摘：詳細化）。
          いずれもセクション背景色（注意喚起トーン）で表現し、3種類あった黄色マーカー装飾は廃止する（R7・v2から継続）。</div>
          <div class="subblocks">
            <div class="subblock tone-a">
              <div class="st">サブブロックA：ブランドレビュー依頼の前にご確認ください</div>
              <div class="sd">レビュー依頼の対象範囲・提出物・所要日数の目安・依頼窓口（マーケへの業務依頼フォーム）を明記する本文ブロック。</div>
              {ai1box("ブランドレビュー依頼の前にご確認ください")}
            </div>
            <div class="subblock tone-b">
              <div class="st">サブブロックB：注意事項（社外秘）</div>
              <div class="sd">社外への転載・二次利用に関する禁止事項、社外秘資料の取扱いルールを明記する本文ブロック。トーンをA（案内）と区別するため背景色を変える。</div>
            </div>
          </div>
          <div class="note">AI-1の設置注記: ブランドの申請/チェック文脈に実装要素として配置（トップのT-4と整合）。サブブロックAに集約し、
          「ロゴ掲載申請」セクション（SECTION 5）からも同じAI-1導線を参照する形にして重複実装を避ける。</div>
        </div>''')(),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 4</div>
              <div><span class="part-tag">[テキスト web part]（既存資料へのリンク・4カテゴリ見出し＋箇条書き）</span></div>
              <div class="title">ガイドラインライブラリ（B-1/B-2・v3で全面再設計）</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span><span class="badge ok">B-1</span><span class="badge ok">B-2</span><span class="badge rule">R5</span></div>
          </div>
          <div class="desc"><b>新規ドキュメントライブラリは作らない</b>（既存ライブラリはルール下で整理済み・変更不可）。目的軸4カテゴリの見出し＋各カテゴリの
          既存資料へのリンク箇条書きを、<b>テキスト web part</b>で表現する（B-2・v2の「ドキュメントライブラリ web part新設」案は撤回）。
          情報量が多いため<b>ページ分割はせず</b>、「このセクションを折りたたみ可能にする」（既定＝折りたたみ）でハブページ内に格納する（B-1・v2のハブ&スポーク案は撤回）。</div>
          {collapsible(
              "既定は折りたたみ・WF上は内容確認のため展開表示",
              linkgroups_block([
                  ("全社ブランド規定", ["AKKODiS ブランドガイドライン V3（既存資料）", "AKKODiS新ビジョンガイドライン（既存資料）"]),
                  ("制作物別ガイドライン", ["ビデオガイドライン（既存資料）", "フォトグラフィーガイドライン（既存資料）", "AI生成画像の使用ガイドライン（既存資料）"]),
                  ("表記・用語", ["文字表記ガイドライン（既存資料）", "グローバル用語集（既存資料）", "ブランドボイスガイドライン（既存資料）"]),
                  ("Akkodis Intelligenceブランド", ["ブランドガイドライン（既存資料）", "Q&A（既存資料）"]),
              ])
          )}
          <div class="note">表示名は「AKKODiS」表記に統一する（B-4・列ではなくリンクラベルの表記統一として反映）。既存ライブラリの中身をそのまま一覧表示したい場合のみ、
          新規作成ではなく<b>既存ライブラリを参照する「ドキュメント ライブラリ web part」</b>を代替手段として使用可（V7）。
          登録前提としてガイドライン10本の重複・旧版の<b>棚卸し</b>を運用タスクとして実施する（WFのコンポーネントではなく運用作業）。</div>
        </div>''')(),
        cols2_block(
            5, "ロゴ掲載申請（手続き・v3でAI-1組込）", "[テキスト web part]（番号付きリスト）＋[ドキュメント ライブラリ web part]（同意書テンプレ・既存）",
            ("①先方企業のロゴを当社サイトに掲載する場合",
             ["AI-1でPPT/画像を事前セルフチェック（サブブロックA参照）", "Marketing部依頼フォーム提出", "Teams連絡",
              "Salesforceサービスリクエストから押印申請", "回収", "送付", "HP掲載"],
             "同意書テンプレ（先方ロゴ掲載用・既存ライブラリ）"),
            ("②当社社名・ロゴを他社サイト等に掲載する場合",
             ['<span class="tbd">（マーケ部確定待ち・例）</span>手順は現物確認の上マーケ部が確定',
              "構造要件のみ規定：①と同様にフォーム起点の番号付き手順として構造化する（具体手順は創作しない）"],
             "同意書テンプレ（当社ロゴ掲載用・既存ライブラリ）"),
            note="2カラムはSharePointの<b>セクション内2列レイアウト</b>（3列以内・標準機能）で実現。手続き自体はv2から変更なし。"
            "①の手順冒頭にAI-1セルフチェックのステップを追加し、SECTION3のAI-1導線と同一である旨を明記（重複実装を避ける）。",
            badges=[("B-3", "ok"), ("R5", "rule")],
        ),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 6</div>
              <div><span class="part-tag">[テキスト web part]（既存素材へのリンク・新規ライブラリ不作成）</span></div>
              <div class="title">ロゴ</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span><span class="badge rule">R3</span></div>
          </div>
          <div class="desc">ロゴ素材データそのものは新規ライブラリを作らず、<b>既存のロゴ素材ライブラリ（サイト内・整理済み）</b>へのリンクをテキストweb部品で貼る。
          「ロゴ掲載申請」（手続き・SECTION5）とは別物である旨を明記する。</div>
          {linkgroups_block([
              ("既存ロゴライブラリ（新規作成なし）", ["AKKODiSロゴ（横組み）→既存ライブラリ内フォルダへリンク", "AKKODiSロゴ（縦組み）→既存ライブラリ内フォルダへリンク",
                                          "モノクロ版→既存ライブラリ内フォルダへリンク", "Akkodis Intelligence版→既存ライブラリ内フォルダへリンク"]),
          ], single_col=True)}
          <div class="note">既存ライブラリの中身を直接見せたい場合は、新規作成ではなく既存ライブラリを参照する「ドキュメント ライブラリ web part」を代替として使用可（V7）。前回WFで欠落していたセクション（新設は維持）。</div>
        </div>''')(),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 7</div>
              <div><span class="part-tag">[テキスト web part]（既存素材へのリンク・新規ライブラリ不作成）</span></div>
              <div class="title">アイコン</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span><span class="badge rule">R3</span></div>
          </div>
          <div class="desc">アイコン素材データそのものは新規ライブラリを作らず、<b>既存のアイコン素材ライブラリ</b>へのリンクをテキストweb部品で貼る。
          冒頭の目次から遷移する本セクションはアイコン<b>素材データ</b>であり、ページ内ナビ用のアイコンとは別物（旧B-5との混同は解消済み・v2で削除済みの旧アンカーカードとの混同懸念も本v3ではSection4削除により解消）。</div>
          {linkgroups_block([
              ("既存アイコンライブラリ（新規作成なし）", ["UI用アイコン→既存ライブラリ内フォルダへリンク", "資料用アイコン→既存ライブラリ内フォルダへリンク", "SNS用アイコン→既存ライブラリ内フォルダへリンク"]),
          ], single_col=True)}
        </div>''')(),
        sec(
            8, "フォント", "[テキスト web part]",
            "現状のテキスト構成を維持する。目次から到達可能にするのみで内容変更はなし（v2から変更なし）。",
        ),
        cards_block(
            9, "画像素材（整理・構造化）", "[テキスト web part]＋[クイックリンク web part]（4サブセクション）",
            [
                ("マーケティング部承認済み画像について", "1〜2行要約。詳細は本カードのアンカー先段落へ"),
                ("画像の購入について", "1〜2行要約。詳細は本カードのアンカー先段落へ"),
                ("AI生成画像の使用について", "1〜2行要約。詳細は本カードのアンカー先段落へ"),
                ("使用禁止の画像について", "1〜2行要約。詳細は本カードのアンカー先段落へ"),
            ],
            note="4カードは3列グリッドのため4枚目が自動で次の行に折り返る（V1・4列表示ではない）。内容はv2から変更なし。",
            badges=[("R4", "rule")],
        ),
        sec(
            10, "各種テンプレート", "[強調表示されたコンテンツ web part]",
            "コンテンツタイプ別（PowerPoint／Word／メール署名／Forms／バーチャル背景）。絞り込み文言はR2の定型表現に統一：<br/>"
            "「<b>[強調表示されたコンテンツ web part] で、事前作成ビュー／Highlighted contentのType指定／管理プロパティ指定により絞り込み"
            "（テナントの検索スキーマ・クロールドプロパティのマッピングに依存・要テナント確認）</b>」。"
            "「ワンクリックで動的に多条件フィルタできる」という誤った表現は使わない。",
            note="【v3】パーツ名を正式名称「強調表示されたコンテンツ web part」に修正（v2の「ハイライトされたコンテンツ」は誤記。support.microsoft.com記事タイトルで確認・確度=高）。",
            badges=[("R2", "rule"), ("v3改訂", "v3")],
        ),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 11</div>
              <div><span class="part-tag">[テキスト web part]（見出しが長いためクイックリンク不可・箇条書き＋リンク）</span></div>
              <div class="title">SNSガイドライン</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span></div>
          </div>
          <div class="desc">現行3リンクは見出しが長いため、クイックリンク web part ではなく<b>テキスト web part（リンク付き箇条書き）</b>で描く（V5：
          クイックリンクの見出し文字数上限は一次情報に数値記載なし・要確認のためリスクを取らない設計）。</div>
          {linkgroups_block([
              ("SNSガイドライン", ["Facebook投稿におけるブランド運用ガイドライン（画像・トーン&マナー）", "X（旧Twitter）投稿時のロゴ・ハッシュタグ運用ルール", "LinkedIn企業ページのビジュアルガイドライン"]),
          ], single_col=True)}
        </div>''')(),
        future_block(
            "AI検索の高度ファセット／Copilot要約等",
            "L1範囲外のためMVPワイヤー本体には含めない。ファセット・リファイナー等の高度な絞り込みはL3（PnP Modern Search等・カスタムコード許可が前提）。（v2から変更なし）",
        ),
        changes_box([
            "【v3】ブランドの1/2・2/2を<b>1枚のWFに統合</b>した。",
            "【v3】旧「Section4 ページ内アンカーカード」（上部目次と重複）を削除した。",
            "【v3】B-1（情報過多）はハブ&スポーク（ページ分割）案を撤回し、「このセクションを折りたたみ可能にする」（既定＝折りたたみ）でハブページ内に格納する設計に変更した。",
            "【v3】B-2（ガイドライン10本）は新規ドキュメントライブラリ作成を撤回し、テキストweb部品（4カテゴリ見出し＋既存資料へのリンク箇条書き）に変更した。",
            "【v3】「ブランドの使用について」をサブブロックA（レビュー依頼前の確認・AI-1導線）／サブブロックB（注意事項・社外秘）に分けて詳細化した。",
            "【v3】「ロゴ」「アイコン」セクションは新規ライブラリを作らず、既存素材へのリンク（テキストweb部品）に変更した。",
            "【v3】「SNSガイドライン」はクイックリンクからテキストweb部品（リンク付き箇条書き）に変更した（見出しが長いため）。",
            "【v3】各種テンプレートのパーツ名を「強調表示されたコンテンツ web part」に修正した（正式名称の誤記訂正）。",
            "【v3】AI-1（ブランドセルフチェッカー）をブランドレビュー依頼・ロゴ掲載申請の文脈に実装要素として組込んだ（将来拡張ではなく実線枠。ライセンス/権限は要確認の小注記付き）。",
        ]),
        footnote("ToBe WF v3 — 2. ブランドライブラリ（1枚に統合）"),
    ]
    return render("brand", 2, TOTAL_PAGES, "2. ブランドライブラリ", blocks)


# ---------------------------------------------------------------------------
# 3. クライアント向け
# ---------------------------------------------------------------------------
def build_client():
    blocks = [
        hero("クライアント向け", "現状維持"),
        toc([
            {"label": "AKKODiS企業向けWebページ", "target": "セミナー／活用事例／Insights／サービス"},
            {"label": "公開資料ライブラリ", "target": "リスト web part（カテゴリ・言語・更新日・ビュー絞り込み）"},
            {"label": "メルマガ配信スケジュール", "target": "配信スケジュール（社外秘）"},
            {"label": "資料を探す", "target": "", "removed": True},
        ]),
        rationale(
            "「資料を探す」アンカー削除の判断根拠（v2から継続・変更なし）",
            "前回WFは「資料を探す（新規）」という4つ目のアンカーを追加していたが、遷移先が明記されておらず社長から指摘を受けた。"
            "精査の結果、意図は①本文検索、または②公開資料ライブラリへのショートカットのいずれかと推測されるが、"
            "①はR1により実装不可、②は既存の目次項目「公開資料ライブラリ」と完全重複する。したがって<b>曖昧な項目として削除</b>した。",
        ),
        cards_block(
            3, "AKKODiS企業向けWebページ", "[クイックリンク web part]",
            [("セミナー", "現状維持"), ("活用事例", "現状維持"), ("AKKODiS Insights", "現状維持"), ("サービスページ", "現状維持")],
            note="4カードは3列グリッドのため4枚目が次の行に自動折り返し（V1）。",
        ),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 4</div>
              <div><span class="part-tag">[リスト web part]（カスタムリスト・ハイパーリンク列）</span></div>
              <div class="title">公開資料ライブラリ（v3で3列以内に再設計）</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span><span class="badge ok">C-1</span><span class="badge ok">C-2</span><span class="badge rule">V1</span></div>
          </div>
          <div class="desc">4カテゴリ（共通／ソリューション／アカデミー／事例資料／お役立ち資料）を<b>列として横に並べない</b>。
          単一の<b>リスト web part</b>内で「行＋列（カテゴリ／言語／更新日）」として表現し、ビューで絞り込む構成にする（社長指摘の反映）。
          資料の実体は外部AKKODiSサイトへのリンクでありファイルではないため、ドキュメントライブラリではなく<b>リスト web part（ハイパーリンク列）</b>を使用する（C-1・V3で正式名称確認済み）。</div>
          <div class="listpart">
            <div class="filterbar">
              <span class="lbl">ビューで絞り込み:</span>
              <span class="chip">共通</span><span class="chip">ソリューション</span><span class="chip">アカデミー</span>
              <span class="chip">事例資料</span><span class="chip">お役立ち資料</span>
            </div>
            <table>
              <tr><th>タイトル（ハイパーリンク列）</th><th>カテゴリ（選択肢列）</th><th>言語（選択肢列）</th><th>更新日（任意列）</th></tr>
              <tr><td class="linkcol">会社案内（日本語版）</td><td class="catcell">共通</td><td>日本語</td><td>—</td></tr>
              <tr><td class="linkcol">会社案内（英語版）</td><td class="catcell">共通</td><td>英語</td><td>—</td></tr>
              <tr><td class="linkcol">ソリューション紹介資料（約11本）</td><td class="catcell">ソリューション</td><td>日本語／英語</td><td>—</td></tr>
              <tr><td class="linkcol">アカデミー紹介資料（約8本）</td><td class="catcell">アカデミー</td><td>日本語／英語</td><td>—</td></tr>
              <tr><td class="linkcol"><span class="tbd">（マーケ部確定待ち・例）</span>事例資料</td><td class="catcell">事例資料（新規追加）</td><td>日本語</td><td>要付与</td></tr>
              <tr><td class="linkcol">お役立ち資料（30本超）</td><td class="catcell">お役立ち資料</td><td>日本語／英語</td><td>任意</td></tr>
            </table>
            <div class="footnote1">
              <b>「ページ内資料の探し方」（旧SECTION5・v3で本枠内の1行注記に簡素化）:</b>
              上部の検索＋この一覧の列（カテゴリ／言語）で絞り込めます。
            </div>
          </div>
          <div class="note">C-1：ドキュメントライブラリ→リスト web part（ハイパーリンク列）に設計変更（R6準拠・変更なし）。C-2：更新日は任意列のまま（付与しない資料があってよい）。
          日本語版／英語版が対の資料は言語列で1エントリに集約。カテゴリに「事例資料」を維持。</div>
        </div>''')(),
        sec(
            5, "メルマガ配信スケジュール", "[クイックリンク web part]",
            "現状維持（社外秘表記も維持）。",
        ),
        changes_box([
            "【v3】「公開資料ライブラリ」を4カテゴリの列並列表示から、<b>単一のリスト web part（行＋カテゴリ/言語/更新日列＋ビュー絞り込みチップ）</b>に再設計し、3列以内（実質1つのリスト枠）で表現した。",
            "【v3】旧SECTION5「ページ内資料の探し方」案内は独立セクションをやめ、公開資料ライブラリの枠内の1行注記に簡素化した（社長指摘：意味不明・過剰）。",
            "【v3】パーツ名を正式名称「リスト web part」に統一した（英語Listの日本語UI表記）。",
            "（v2からの継続）「事例資料」カテゴリを維持。「資料を探す」アンカーは重複のため削除済み。「このページ内資料検索」はR1により実装不可のため削除済み。",
        ]),
        footnote("ToBe WF v3 — 3. クライアント向け"),
    ]
    return render("client", 3, TOTAL_PAGES, "3. クライアント向け", blocks)


# ---------------------------------------------------------------------------
# 4. キャンディデート向け
# ---------------------------------------------------------------------------
def build_candidate():
    blocks = [
        hero("キャンディデート向けページ", "現状維持"),
        toc([
            {"label": "AKKODiS公式HP", "target": "AKKODiS People（文脈内リンクで業務依頼フォームへ）／リクルーティングムービー"},
            {"label": "マスター資料", "target": "会社説明資料（現行踏襲・3チーム連絡先）"},
            {"label": "Social Media", "target": "X／Facebook／LinkedIn／YouTube"},
        ]),
        cards_block(
            2, "AKKODiS公式HP", "[クイックリンク web part]",
            [
                ("AKKODiS People",
                 "現状維持。<br/><span class=\"inline-link\">→ ※SNS等の告知投稿の依頼は業務依頼フォームから</span>"),
                ("リクルーティングムービー", "現状維持"),
            ],
            note="【v3】業務依頼フォームへのリンクは、並列の別カードではなく<b>「AKKODiS People」項目の説明文脈内のリンク</b>として1つ添える形に変更した（社長指摘：完全並列は誤り）。",
            badges=[("v3改訂", "v3")],
        ),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 3</div>
              <div><span class="part-tag">[テキスト web part]（現行踏襲・そのまま掲載）</span></div>
              <div class="title">マスター資料（CA-1・v3で現行文言を確定掲載）</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span><span class="badge ok">CA-1確定・現行踏襲</span></div>
          </div>
          <div class="desc">直リンク化案は不採用とし、<b>現行の文言をそのままテキストweb部品に掲載する</b>（社長指摘：CA-1は現行踏襲。旧「担当確認待ち・暫定」ラベルは撤去し、確定的に描く）。</div>
          <div class="subblocks">
            <div class="subblock">
              <div class="st">現行文言（そのまま掲載）</div>
              <div class="sd">「会社説明資料は各チームで管理しています。必要な場合は各チームまで直接ご連絡ください。」</div>
            </div>
          </div>
          <div class="lib-meta">
            <span>中途採用向け：担当チーム連絡先</span>
            <span>新卒採用向け：担当チーム連絡先</span>
            <span>障がい者採用向け：担当チーム連絡先</span>
          </div>
          <div class="note">3チーム連絡先は現状維持のメタ情報として保持する（直リンク化はしない）。</div>
        </div>''')(),
        cards_block(
            4, "Social Media", "[クイックリンク web part]",
            [("X", "現状維持"), ("Facebook", "現状維持"), ("LinkedIn", "現状維持"), ("YouTube", "現状維持")],
            note="4カードは3列グリッドのため4枚目が次の行に自動折り返し（V1）。",
        ),
        changes_box([
            "【v3】「AKKODiS公式HP」の業務依頼フォームへのリンクを、並列カードから「AKKODiS People」項目の<b>文脈内リンク</b>（※SNS等の告知投稿の依頼は業務依頼フォームから）に変更した。",
            "【v3】CA-1（マスター資料）は「担当確認待ち・暫定」ラベルを撤去し、<b>現行文言をそのままテキストweb部品に確定掲載</b>する設計にした（直リンク化案は不採用）。",
            "（v2からの継続）CA-2（3セクション構成の簡潔さ）は変更していない。",
        ]),
        footnote("ToBe WF v3 — 4. キャンディデート向け"),
    ]
    return render("candidate", 4, TOTAL_PAGES, "4. キャンディデート向け", blocks)


# ---------------------------------------------------------------------------
# 5. グローバルブランドパートナーシップ
# ---------------------------------------------------------------------------
def build_partner():
    blocks = [
        hero("グローバルブランドパートナーシップ", "ヒーロー＋リード文（現状維持）"),
        toc([
            {"label": "Mercedes-AMG PETRONAS Formula One Team", "target": "該当パートナーセクション"},
            {"label": "Akkodis ASP Team", "target": "該当パートナーセクション"},
            {"label": "Stade Toulousain", "target": "該当パートナーセクション"},
        ], note="全ページ共通の目次パターン（G-1）に統一。各アンカーは同一ページ内の該当パートナーセクションへジャンプする。"),
        rationale(
            "サムネイルカード削除の判断根拠（v3・社長指摘の反映）",
            "前回WFは「3パートナーのサムネイルカード」と「ページ内目次」を役割A/Bとして併記していたが、"
            "両者ともリンク先が同じ3セクションで視覚的にも似ており、社長から重複感の指摘を受けた。"
            "本v3では<b>サムネイルカードを削除</b>し、全ページ共通の目次（クイックリンク web part・体裁統一）のみを残す構成に変更した。",
        ),
        sec(
            3, "各パートナーセクション（3ブロック）", "[テキスト web part]（H2見出し）＋既存構成",
            "Mercedes-AMG PETRONAS Formula One Team／Akkodis ASP Team／Stade Toulousainの3ブロック。"
            "概要／共通価値／パートナーシップビデオ（現状維持）。各ブロック先頭の見出しが上部目次のアンカー先になる。",
            badges=[("R4", "rule")],
        ),
        sec(
            4, "Akkodis ASP Teamの画像枠（P-1・修正済み）", "[画像 web part]",
            "他2チームと同じ構造に統一済み（維持・変更なし）。",
        ),
        changes_box([
            "【v3】3パートナーのサムネイルカード（役割A）を削除し、ページ内目次（クイックリンク web part）のみを残した（社長指摘：全ページ体裁統一）。",
            "（v2からの継続）P-1（ASP Teamの画像枠是正）・P-2（ページ内目次新設）は維持している。",
        ]),
        footnote("ToBe WF v3 — 5. グローバルブランドパートナーシップ"),
    ]
    return render("partner", 5, TOTAL_PAGES, "5. グローバルブランドパートナーシップ", blocks)


# ---------------------------------------------------------------------------
# PDF結合（1ページ＝1WF）
# ---------------------------------------------------------------------------
def build_pdf(png_paths, out_name="ToBe_WF_v3_承認用.pdf"):
    imgs = [Image.open(p).convert("RGB") for p in png_paths]
    out_path = os.path.join(OUT, out_name)
    imgs[0].save(out_path, save_all=True, append_images=imgs[1:])
    print("saved", out_path)
    return out_path


if __name__ == "__main__":
    paths = []
    paths.append(build_top())
    paths.append(build_brand())
    paths.append(build_client())
    paths.append(build_candidate())
    paths.append(build_partner())
    build_pdf(paths)
    print("all done")

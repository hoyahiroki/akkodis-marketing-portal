#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ToBe ワイヤーフレーム v5 生成スクリプト（判読性最優先・大判レンダリング）
- v4（build_tobe_wf_v4.py）を土台に、社長4回目差戻し＋一次情報調査
  （sharepoint-verify-v5.md V14〜V17）を反映した改訂版。**指示された変更のみ**。
- v4→v5の主な変更点（詳細は wf-v5-directives.md）:
  1) ブランド「各種テンプレート」：v4で採用した「強調表示されたコンテンツ web part」×5を
     全廃し、「テキスト web part」1つに変更（V17：強調表示されたコンテンツはサイト内の
     ファイル/ページを動的クエリで一覧表示するパーツであり、リンク先が1〜6個の外部/内部
     リンク集には使えない。社長指摘が正しく、v4時点の設計は誤りだった＝撤回）。
     5区分を見出し＋リンク箇条書きで表現し、実際のリンク数（6/1/1/1/2個）を反映する。
  2) クライアント「公開資料ライブラリ」：v4までの「リスト web part」から「テキスト web part」
     を主案に変更（V15：SharePoint agent＝AI-2は「Agents currently don't use data from
     Lists.」と公式に明記されており、リストのデータを回答根拠に一切使わない。AI-2に拾わせ
     たい資料リンクはページ本文＝テキストweb partに書く必要がある）。カテゴリ見出し＋リンク
     箇条書き、お役立ち資料はセクション折りたたみで格納。トレードオフ（カテゴリ/言語の動的な
     絞り込み・並べ替えを失う）を明記し、副案（点線枠）として「リスト併用案」（V14：リスト
     アイテムは既定でMicrosoft Searchにインデックスされ人間の検索には有効／但しAI-2の回答
     根拠にはならない）を併記する。
  3) 全ページ共通：AI-2注記帯に「リストのデータは対象外（ページ本文・ドキュメントライブラリ
     が対象）」を1行追記（V15）。
  4) トップ／キャンディデート／パートナーは変更なし。ただしキャンディデートには
     「AKKODiS Peopleのテキスト化・Social Media 4列はv4で対応済み」の注記を目立つ形で追加
     （社長が同一指摘を再送されたため、対応済みであることが一目で分かるようにする）。
"""
import os
from playwright.sync_api import sync_playwright
from PIL import Image

OUT = "/workspace/akkodis-marketing-portal/deliverables/1784786903168-17720/wf_v5"
os.makedirs(OUT, exist_ok=True)

CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

WIDTH = 1160  # 本文幅（v3から変更なし）
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
  background:#eaf3ec; border-bottom:3px solid #2f6b3f; padding:8px 26px 10px;
  font-size:12.5px; color:#1c4a28;
}
.ai2band .row1 { display:flex; align-items:center; gap:12px; flex-wrap:wrap; }
.ai2band .tag {
  background:#2f6b3f; color:#fff; padding:3px 10px; border-radius:3px;
  font-size:11px; font-weight:bold; white-space:nowrap;
}
.ai2band .body { color:#1c4a28; }
.ai2band .lic { font-size:10.5px; color:#5a7d63; margin-left:auto; }
.ai2band .row2 {
  margin-top:6px; display:flex; align-items:center; gap:10px; flex-wrap:wrap;
  background:#fff3de; border:1.5px solid #c58a2e; border-radius:3px; padding:5px 12px;
}
.ai2band .row2 .tag2 {
  background:#8a5518; color:#fff; padding:2px 9px; border-radius:3px; font-size:10.5px; font-weight:bold; white-space:nowrap;
}
.ai2band .row2 .body2 { color:#5a3a10; font-size:12px; font-weight:bold; }

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
.badge.v4 { background:#e3edfb; border-color:#2f5f9e; color:#123a6b; font-weight:bold; }
.badge.v5 { background:#f7e3fb; border-color:#8a3b9e; color:#5a1f6b; font-weight:bold; }
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

/* T-4（v4）: 縦積み3ステップ（①→②→③を上から下へ全幅で配置。3カラム横並びをやめる） */
.flowV { display:flex; flex-direction:column; margin-top:10px; }
.flowV .stepRow { display:flex; gap:16px; align-items:stretch; }
.flowV .stepnum {
  flex:0 0 30px; width:30px; height:30px; border-radius:50%;
  background:#555; color:#fff; font-size:14px; font-weight:bold;
  display:flex; align-items:center; justify-content:center;
}
.flowV .stepbox { flex:1; }
.flowV .arrowDown { margin-left:14px; padding:4px 0; font-size:20px; color:#8a8f96; line-height:1; }
.flowV .guidebox { border:2px solid #8a8f96; background:#fff; padding:16px 20px; }
.flowV .guidebox .t { font-size:14.5px; font-weight:bold; color:#1e2126; margin-bottom:6px; }
.flowV .guidebox .d { font-size:13px; color:#555; line-height:1.7; }
.flowV .formbox {
  border:2px solid #2b3a55; background:#eef1f6; padding:16px 20px;
  display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px;
}

/* カード群（V1: セクション列レイアウトは3列以内。クイックリンクweb part自体は
   V8再検証によりレスポンシブに4列以上が可能＝.cards.cols4 で表現。既定は3列） */
.cards { display:grid; grid-template-columns: repeat(3, 1fr); gap:14px; margin-top:6px; }
.cards.cols4 { grid-template-columns: repeat(4, 1fr); }
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

/* セクション単位の背景色トーン（v4: サブブロック単位の色分けはセクション分割で実現・V9） */
.sec.tone-a { background:#fff4e0; border-color:#c58a2e; }
.sec.tone-b { background:#fdeceb; border-color:#c0392b; }

/* 強調表示されたコンテンツ web part ×5（v4: カテゴリごとに個別web partとして配置・V10） */
.hc5 { display:grid; grid-template-columns: repeat(3, 1fr); gap:14px; margin-top:8px; }
.hc5 .hcitem { border:2px solid #5c7ba8; background:#eef3fa; padding:14px 16px; }
.hc5 .hcitem .tag {
  display:inline-block; font-size:10.5px; font-weight:bold; color:#fff;
  background:#3b5b8c; padding:2px 8px; border-radius:3px; margin-bottom:7px;
}
.hc5 .hcitem .t { font-size:13.5px; font-weight:bold; color:#1c3a5e; margin-bottom:6px; }
.hc5 .hcitem .d { font-size:12px; color:#33455c; line-height:1.65; }

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

/* v5: 対応済み再周知の強調ボックス（社長が同一指摘を再送した箇所に使用） */
.resolved {
  margin:20px 24px; border:3px solid #1c5f3a; background:#e8f7ec; padding:14px 20px;
}
.resolved .tag {
  display:inline-block; font-size:12px; font-weight:bold; color:#fff; background:#1c5f3a;
  padding:3px 10px; border-radius:4px; margin-bottom:8px;
}
.resolved .t { font-size:15px; font-weight:bold; color:#124a2e; margin-bottom:6px; }
.resolved .d { font-size:13px; color:#1c4a33; line-height:1.8; }

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
      <div class="name">ToBe ワイヤーフレーム v5　｜　{name}　（{idx}/{total}）</div>
      <div class="sub">{sub or "AKKODiSマーケティングポータル・承認用WF（社長4回目差戻し反映）"}</div>
    </div>'''


def searchband():
    return '''<div class="searchband">
      <span class="tag">システム標準・編集不可</span>
      <span>サイトヘッダー：Microsoft Search（既定表示・位置固定・全ページ共通）</span>
      <span class="desc">※本文コンポーネントではない。ページ本文には検索ボックスを配置しない（R1）</span>
    </div>'''


def ai2band():
    return '''<div class="ai2band">
      <div class="row1">
        <span class="tag">実装要素・全ページ共通</span>
        <span class="body">サイト共通エージェント：<b>AI-2</b>（SharePoint agent・ヘッダーの「エージェント」アイコンから起動・利用者本人のアクセス権範囲でのみ回答・出典/引用リンク付き）</span>
        <span class="lic">※要 Microsoft 365 Copilotライセンス or Pay-as-you-go課金（要テナント管理者確認）</span>
      </div>
      <div class="row2">
        <span class="tag2">v5追記・V15</span>
        <span class="body2">リストのデータは対象外（ページ本文・ドキュメントライブラリが対象）。公式明記："Agents currently don't use data from Lists."</span>
      </div>
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


def cards_block(num, title, part, cards, note=None, badges=None, new=False, cols=3):
    badges_html = "".join(f'<span class="badge {b[1]}">{b[0]}</span>' for b in (badges or []))
    cs = "".join(
        f'<div class="card"><div class="icon"></div><div class="t">{c[0]}</div><div class="d">{c[1]}</div></div>'
        for c in cards
    )
    note_html = f'<div class="note">{note}</div>' if note else ""
    klass = "sec new" if new else "sec"
    cards_cls = "cards cols4" if cols == 4 else "cards"
    return f'''<div class="{klass}">
      <div class="sechead">
        <div>
          <div class="num">SECTION {num}</div>
          <div><span class="part-tag">{part}</span></div>
          <div class="title">{title}</div>
        </div>
        <div class="badges">{badges_html}</div>
      </div>
      <div class="{cards_cls}">{cs}</div>
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


def hc5_block(num, title, items, note=None, badges=None):
    """v4: 「強調表示されたコンテンツ web part」をカテゴリ数ぶん個別配置する（V10：グループ表示機能が無いため）。
    items = [(見出し, フィルタ説明), ...]。セクション列は3列以内のため、4件目以降はグリッドが自動で次の行に折り返す。"""
    badges_html = "".join(f'<span class="badge {b[1]}">{b[0]}</span>' for b in (badges or []))
    its = "".join(
        f'<div class="hcitem"><span class="tag">[強調表示されたコンテンツ web part]</span>'
        f'<div class="t">{t}</div><div class="d">{d}</div></div>'
        for t, d in items
    )
    note_html = f'<div class="note">{note}</div>' if note else ""
    return f'''<div class="sec">
      <div class="sechead">
        <div>
          <div class="num">SECTION {num}</div>
          <div><span class="part-tag">[強調表示されたコンテンツ web part] × {len(items)}（カテゴリごとに個別配置）</span></div>
          <div class="title">{title}</div>
        </div>
        <div class="badges">{badges_html}</div>
      </div>
      <div class="hc5">{its}</div>
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


def resolved_notice(title, desc):
    """v5: 社長が同一指摘を再送した箇所に、既に対応済みであることを一目で示す強調ボックス。"""
    return f'''<div class="resolved">
      <span class="tag">v4で対応済み・再指摘への回答（v5でも変更なし）</span>
      <div class="t">{title}</div>
      <div class="d">{desc}</div>
    </div>'''


def altbox(tag, title, desc, extra=""):
    """v5: 主案に対する副案を点線枠で併記するための汎用ボックス（.future の点線スタイルを流用）。
    社長が主案/副案を選べる形にするため、将来拡張(L2/L3)とは別に用意する。"""
    return f'''<div class="future">
      <span class="tag">{tag}</span>
      <div class="t">{title}</div>
      <div class="d">{desc}</div>
      {extra}
    </div>'''


def changes_box(lines, head="このページの変更点（v5・社長4回目差戻し反映）", pin="v5変更点"):
    lis = "".join(f"<li>{l}</li>" for l in lines)
    return f'''<div class="changes">
      <div class="h"><span class="pin">{pin}</span>{head}</div>
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
            {"label": "②マーケへの業務依頼", "target": "AI-1セルフチェック→手順ガイド→業務依頼フォーム（縦積み3ステップ）"},
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
            note="各カード直下に1行の説明文欄（T-3対応・変更なし）。【v4】4列に戻した：V8再検証によりクイックリンクweb part自体は"
            "レスポンシブに4個横並び表示が可能と確認。3列上限はセクションの列レイアウト（強調表示されたコンテンツ/リスト等）にのみ適用される。",
            badges=[("T-3", "ok"), ("v4改訂", "v4"), ("V8", "rule")],
            cols=4,
        ),
        (lambda: f'''<div class="sec new">
          <div class="sechead">
            <div>
              <div class="num">SECTION 4</div>
              <div><span class="part-tag">①[Embed web part]　②[テキスト web part]　③[Call to action web part]</span></div>
              <div class="title">マーケへの業務依頼（v4：縦積み3ステップに再設計）</div>
            </div>
            <div class="badges"><span class="badge v4">v4改訂</span><span class="badge ok">T-4</span><span class="badge rule">G-3</span></div>
          </div>
          <div class="desc">依頼前にAI-1で自己チェック→手順ガイドを確認→業務依頼フォーム送信、という<b>3ステップの導線</b>を、
          <b>①→②→③を上から下へ全幅で配置する縦積みレイアウト</b>に変更する（社長提案・v4。3カラム横並びはやめ、
          各ステップが全幅で内容を余裕を持って表示できるようにする）。</div>
          <div class="flowV">
            <div class="stepRow">
              <div class="stepnum">①</div>
              <div class="stepbox">{ai1box("マーケへの業務依頼（依頼前セルフチェック）")}</div>
            </div>
            <div class="arrowDown">↓</div>
            <div class="stepRow">
              <div class="stepnum">②</div>
              <div class="stepbox guidebox">
                <div class="t">[テキスト web part] 手順ガイド</div>
                <div class="d">「依頼前にご確認ください：ブランドレビューの手順」。フォームより小さい副導線・確認用リンクとして配置（v2から維持）。</div>
              </div>
            </div>
            <div class="arrowDown">↓</div>
            <div class="stepRow">
              <div class="stepnum">③</div>
              <div class="stepbox formbox">
                <div class="cta-primary">[Call to action web part]<br/>マーケへの業務依頼フォーム</div>
              </div>
            </div>
          </div>
          <div class="cta-relation">導線の関係説明: 「①AI-1で依頼前に自己チェック（気づきの提示）→②手順ガイドで進め方を確認→③フォームから正式に依頼」。
          横並びの3枚のカードではなく、上から下へ進む<b>一本の縦積み依頼フロー</b>として描く。</div>
          <div class="note">【v4】旧v3は3カラム横並び（flow3）だったが、社長提案により<b>縦積み・全幅3ステップ</b>に変更した。
          AI-1／手順ガイドはあくまで依頼前の確認・自己チェック導線であり、最終的な依頼窓口はフォーム1つに統一されている（G-3・変更なし）。</div>
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
            "【v5】本ページは<b>変更なし</b>（v5変更指示：トップ／キャンディデート／パートナーは変更なし）。全ページ共通のAI-2注記帯にのみ"
            "「リストのデータは対象外」の1行を追記（本ページはそもそもリストweb partを使用していないため実質影響なし）。",
            "【v4】4カテゴリカードを<b>4列表示に戻した</b>（V8再検証：クイックリンクweb part自体はレスポンシブに4個以上横並び可能。"
            "前回の「全パーツ3列化」は過剰修正だったと判明。3列上限はセクションの列レイアウト＝強調表示されたコンテンツ/リスト等をセクションの列に並べる場合のみ）。",
            "【v4】SECTION4「マーケへの業務依頼」の3ステップ導線を<b>3カラム横並びから縦積み（上から下へ全幅）に変更</b>した（社長提案）。",
            "（v3までの継続）AI-1（ブランドセルフチェッカー）を実装要素として組込み。全ページ共通でAI-2（SharePoint agent）実装注記帯を配置。目次アンカー新設。モバイル対応（T-1）はWFから削除済み。",
        ]),
        footnote("ToBe WF v5 — 1. トップ（ホーム） — v5変更指示：本ページは変更なし（sharepoint-verify-v4.md V8〜V13／v5はV14〜V17）"),
    ]
    return render("top", 1, TOTAL_PAGES, "1. トップ（ホーム）", blocks)


# ---------------------------------------------------------------------------
# 2. ブランドライブラリ（v3: 1枚に統合）
# ---------------------------------------------------------------------------
def build_brand():
    blocks = [
        hero("ブランドライブラリ", "現状維持"),
        toc([
            {"label": "ブランドの使用について", "target": "独立2セクション（1/2 レビュー依頼前確認・AI-1導線／2/2 注意事項・社外秘。背景色による色分けはセクション分割で実現）"},
            {"label": "ガイドライン一覧", "target": "テキストwebパーツ・4カテゴリ見出し＋リンク箇条書き（折りたたみ格納）"},
            {"label": "ロゴ掲載申請", "target": "2ケース手続き（AI-1組込）"},
            {"label": "ロゴ", "target": "既存ロゴ素材へのリンク"},
            {"label": "アイコン", "target": "既存アイコン素材へのリンク"},
            {"label": "フォント", "target": "フォント"},
            {"label": "画像素材", "target": "テキストweb part（4見出し＋リンク箇条書き。クイックリンク併用は廃止）"},
            {"label": "各種テンプレート", "target": "【v5】テキスト web part 1つ（5見出し＋リンク箇条書き。強調表示されたコンテンツ web part×5は全廃）"},
            {"label": "SNSガイドライン", "target": "テキストwebパーツ・リンク箇条書き"},
        ], note="9アンカー（項目数は変更なし）。旧「Section4 ページ内アンカーカード」は上部目次と機能重複のため削除済み（v3）。"),
        (lambda: f'''<div class="sec new tone-a">
          <div class="sechead">
            <div>
              <div class="num">SECTION 3</div>
              <div><span class="part-tag">[テキスト web part]（独立セクション・セクション背景色＝注意喚起トーンA）</span></div>
              <div class="title">ブランドの使用について（1/2）：ブランドレビュー依頼の前にご確認ください</div>
            </div>
            <div class="badges"><span class="badge v4">v4改訂</span><span class="badge rule">V9</span><span class="badge rule">R7</span></div>
          </div>
          <div class="desc">レビュー依頼の対象範囲・提出物・所要日数の目安・依頼窓口（マーケへの業務依頼フォーム）を明記する本文ブロック。</div>
          {ai1box("ブランドレビュー依頼の前にご確認ください")}
          <div class="note">【v4】旧v3では本ブロックと「注意事項（社外秘）」を1セクション内のサブブロック（左ボーダー色分け）として同居させていたが、
          V9検証の結果「背景色（Section background shading）はセクション単位でのみ設定可能。テキストweb part内の段落単位で個別の背景色を付ける機能は
          公式ドキュメントに記載がなくUI上も存在しない」と判明。<b>サブブロック単位の色分けはセクション分割で実現する</b>ため、本ブロックを独立セクションに分割した。
          AI-1の設置注記: ブランドの申請/チェック文脈に実装要素として配置（トップのT-4と整合）。「ロゴ掲載申請」セクション（SECTION 6）からも同じAI-1導線を参照する形にして重複実装を避ける。</div>
        </div>''')(),
        (lambda: f'''<div class="sec tone-b">
          <div class="sechead">
            <div>
              <div class="num">SECTION 4</div>
              <div><span class="part-tag">[テキスト web part]（独立セクション・セクション背景色＝注意喚起トーンB）</span></div>
              <div class="title">ブランドの使用について（2/2）：注意事項（社外秘）</div>
            </div>
            <div class="badges"><span class="badge v4">v4改訂</span><span class="badge rule">V9</span><span class="badge rule">R7</span></div>
          </div>
          <div class="desc">社外への転載・二次利用に関する禁止事項、社外秘資料の取扱いルールを明記する本文ブロック。
          トーンA（案内・SECTION 3）と区別するため<b>独立セクションの背景色</b>を変える。</div>
          <div class="note">サブブロック単位の色分けはセクション分割で実現する（V9）。3種類あった黄色マーカー装飾は廃止済み（R7・v3から継続）。</div>
        </div>''')(),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 5</div>
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
            6, "ロゴ掲載申請（手続き・v3でAI-1組込）", "[テキスト web part]（番号付きリスト）＋[ドキュメント ライブラリ web part]（同意書テンプレ・既存）",
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
              <div class="num">SECTION 7</div>
              <div><span class="part-tag">[テキスト web part]（既存素材へのリンク・新規ライブラリ不作成）</span></div>
              <div class="title">ロゴ</div>
            </div>
            <div class="badges"><span class="badge v3">v3改訂</span><span class="badge rule">R3</span></div>
          </div>
          <div class="desc">ロゴ素材データそのものは新規ライブラリを作らず、<b>既存のロゴ素材ライブラリ（サイト内・整理済み）</b>へのリンクをテキストweb部品で貼る。
          「ロゴ掲載申請」（手続き・SECTION6）とは別物である旨を明記する。</div>
          {linkgroups_block([
              ("既存ロゴライブラリ（新規作成なし）", ["AKKODiSロゴ（横組み）→既存ライブラリ内フォルダへリンク", "AKKODiSロゴ（縦組み）→既存ライブラリ内フォルダへリンク",
                                          "モノクロ版→既存ライブラリ内フォルダへリンク", "Akkodis Intelligence版→既存ライブラリ内フォルダへリンク"]),
          ], single_col=True)}
          <div class="note">既存ライブラリの中身を直接見せたい場合は、新規作成ではなく既存ライブラリを参照する「ドキュメント ライブラリ web part」を代替として使用可（V7）。前回WFで欠落していたセクション（新設は維持）。</div>
        </div>''')(),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 8</div>
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
            9, "フォント", "[テキスト web part]",
            "現状のテキスト構成を維持する。目次から到達可能にするのみで内容変更はなし（v2から変更なし）。",
        ),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 10</div>
              <div><span class="part-tag">[テキスト web part]（4見出し＋各リンクを箇条書き。クイックリンク併用は廃止）</span></div>
              <div class="title">画像素材（v4：テキストweb partのみに簡素化）</div>
            </div>
            <div class="badges"><span class="badge v4">v4改訂</span><span class="badge rule">V11</span></div>
          </div>
          <div class="desc">【v4】旧v3はテキストweb part＋クイックリンクweb partの併用（カード風の見出し＋アンカー先段落）だったが、
          <b>クイックリンク併用をやめ、テキストweb part単体</b>に簡素化する（社長指摘）。1つのテキストweb part内に
          4見出し＋各見出し配下のリンクを箇条書きで表現する。</div>
          {linkgroups_block([
              ("マーケティング部承認済み画像について", ["承認済み画像バンク（既存ライブラリ）へのリンク", "利用申請はマーケへの業務依頼フォームから"]),
              ("画像の購入について", ["推奨ストックフォトサービス一覧（既存資料）", "購入時の申請フロー（既存資料）"]),
              ("AI生成画像の使用について", ["AI生成画像の使用ガイドライン（既存資料）", "利用可否チェックリスト（既存資料）"]),
              ("使用禁止の画像について", ["使用禁止の例一覧（既存資料）", "著作権に関する注意事項（既存資料）"]),
          ])}
          <div class="note">4見出しは1つのテキストweb part内の段落構成（2列はWF上の便宜的なレイアウトであり、SharePointのセクション列ではない）。内容（4カテゴリ）自体はv3から変更なし。</div>
        </div>''')(),
        (lambda: f'''<div class="sec new">
          <div class="sechead">
            <div>
              <div class="num">SECTION 11</div>
              <div><span class="part-tag">[テキスト web part] 1つ（5見出し＋リンク箇条書き）</span></div>
              <div class="title">各種テンプレート（v5：テキストweb partに全面変更・強調表示コンテンツは全廃）</div>
            </div>
            <div class="badges"><span class="badge v5">v5改訂</span><span class="badge removed">強調表示Web Part×5→全廃</span><span class="badge rule">V17</span></div>
          </div>
          <div class="desc">【v5】v4で採用した「強調表示されたコンテンツ web part」×5は<b>全廃</b>し、<b>テキスト web part 1つ</b>に変更する
          （社長指摘が正しく、v4時点の設計は誤りだったため撤回）。既存の5区分を見出し＋リンク箇条書きで表現し、<b>実際のリンク数</b>を反映する。</div>
          {linkgroups_block([
              ("Microsoft Officeを用いた資料作成ガイドライン、テンプレート（リンク6個）",
               [f'<span class="tbd">（マーケ部確定待ち・例）</span>既存資料へのリンク {i}' for i in range(1, 7)]),
              ("WEB会議用バーチャル背景（リンク1個）",
               ['<span class="tbd">（マーケ部確定待ち・例）</span>既存資料へのリンク 1']),
              ("メール署名テンプレート（リンク1個）",
               ['<span class="tbd">（マーケ部確定待ち・例）</span>既存資料へのリンク 1']),
              ("社員の写真掲載に関する許諾同意書（リンク1個）",
               ['<span class="tbd">（マーケ部確定待ち・例）</span>既存資料へのリンク 1']),
              ("Formsテンプレート（リンク2個）",
               [f'<span class="tbd">（マーケ部確定待ち・例）</span>既存資料へのリンク {i}' for i in range(1, 3)]),
          ], single_col=True)}
          <div class="note"><b>不採用の理由（V17）:</b> 強調表示されたコンテンツ web partは「サイト内のファイル/ページ/ニュース等を動的クエリで
          一覧表示」するパーツであり、ソースとして選べるのは This site／A document library on this site／This site collection／
          Select sites（最大30サイト）等、<b>サイト・ライブラリ単位の動的クエリのみ</b>。<b>個別の外部/内部リンクを1件ずつ指定して
          リンク先が1〜6個のリンク集を作る用途には使えない</b>（一覧に該当の選択肢が存在しない）。社長指摘の通り、v4時点でこのパーツを
          採用した設計は誤りだった。少数固定のリンク集は<b>テキスト web part</b>（または見た目を整えたい場合はQuick Links web part＝
          クイックリンクweb part）で表現するのが正しい（V17）。</div>
        </div>''')(),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 12</div>
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
            "【v5】「各種テンプレート」を<b>5つの「強調表示されたコンテンツ web part」から「テキスト web part」1つに全面変更</b>した"
            "（V17検証：強調表示されたコンテンツはサイト/ライブラリ単位の動的一覧用であり、リンク先が1〜6個のリンク集には使えないと確認。"
            "社長指摘が正しく、v4時点の設計を撤回。実際のリンク数＝6/1/1/1/2個を反映した見出し＋箇条書きに変更）。",
            "【v4】「ブランドの使用について」のサブブロックA/Bを<b>別セクションに分割</b>した（SECTION3／SECTION4。V9検証：背景色はセクション単位でのみ設定可能なため、"
            "サブブロック単位の色分けはセクション分割で実現する。各サブブロックに独立したセクション背景色を設定）。",
            "【v4】「画像素材」を<b>テキストweb partのみに簡素化</b>した（クイックリンク併用をやめ、テキストパーツ内に4見出し＋各リンクの箇条書きで表現）。",
            "（v3までの継続）ブランドの1/2・2/2を1枚のWFに統合。旧「Section4 ページ内アンカーカード」削除。B-1は折りたたみ格納、B-2は新規ライブラリ不作成に変更済み。"
            "「ロゴ」「アイコン」は既存素材へのリンクに変更済み。「SNSガイドライン」はテキストweb部品に変更済み。AI-1をブランドレビュー依頼・ロゴ掲載申請の文脈に実装要素として組込み済み。"
            "（セクション番号は本v4の分割・再構成に伴い3番以降が1つずつ繰り下がっている）",
        ]),
        footnote("ToBe WF v5 — 2. ブランドライブラリ（1枚に統合）— SharePoint UI再検証 v5（V17。旧V9・V10・V11も反映継続）"),
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
            {"label": "公開資料ライブラリ", "target": "【v5主案】テキスト web part（カテゴリ見出し＋リンク箇条書き）／【副案・点線枠】リスト併用"},
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
            note="【v4】4列に戻した：V8再検証によりクイックリンクweb part自体はレスポンシブに4個横並び表示が可能と確認。"
            "3列上限はセクションの列レイアウト（強調表示されたコンテンツ/リスト等をセクションの列に並べる場合）にのみ適用される。",
            badges=[("v4改訂", "v4"), ("V8", "rule")],
            cols=4,
        ),
        (lambda: f'''<div class="sec new">
          <div class="sechead">
            <div>
              <div class="num">SECTION 4</div>
              <div><span class="part-tag">[テキスト web part]（カテゴリ見出し＋リンク箇条書き・主案）</span></div>
              <div class="title">公開資料ライブラリ（v5：テキストweb part主案に変更）</div>
            </div>
            <div class="badges"><span class="badge v5">v5改訂</span><span class="badge ok">C-1</span><span class="badge ok">C-2</span><span class="badge rule">V15</span></div>
          </div>
          <div class="desc">【v5】v4までの<b>リスト web part</b>から<b>テキスト web part</b>を<b>主案</b>に変更する。理由：<b>SharePoint agent（AI-2）は
          リストのデータを回答根拠に一切使わない</b>（V15：公式に "Agents currently don't use data from Lists." と明記。対象はサイト・
          ドキュメントライブラリ・フォルダ・ファイルのみ）。AI-1・AI-2の実装がマスト要件のリニューアルである以上、資料リンクはAI-2が拾える
          <b>ページ本文（テキストweb part）</b>に置く必要がある。</div>
          {linkgroups_block([
              ("共通", ["会社案内（日本語版）", "会社案内（英語版）"]),
              ("ソリューション", ["ソリューション紹介資料（約11本・日本語／英語）"]),
              ("アカデミー", ["アカデミー紹介資料（約8本・日本語／英語）"]),
              ("事例資料", ['<span class="tbd">（マーケ部確定待ち・例）</span>事例資料'] ),
          ], single_col=True)}
          {collapsible(
              "既定は折りたたみ・WF上は内容確認のため展開表示",
              linkgroups_block([
                  ("お役立ち資料（30本超）", [
                      '<span class="tbd">（マーケ部確定待ち・例）</span>お役立ち資料 例①（日本語）',
                      '<span class="tbd">（マーケ部確定待ち・例）</span>お役立ち資料 例②（英語）',
                      '…他 30本超（マーケ部確定待ち・全件の棚卸しが前提）',
                  ]),
              ], single_col=True)
          )}
          <div class="note"><b>「ページ内資料の探し方」（旧SECTION5・枠内の1行注記）:</b> 上部の検索（サイト内検索・人間向け）＋本ページの
          カテゴリ見出しで資料を探せます。AI-2に質問する場合は、この本文に書かれたリンクが回答の根拠になります（列フィルターでの絞り込みは
          テキストweb part化に伴い提供しません＝下記トレードオフ参照）。</div>
          <div class="note"><b>トレードオフ:</b> テキストweb part化により、リストが持っていた<b>カテゴリ/言語の動的な絞り込み・並べ替え</b>
          （ビュー機能）は失われる。一方でAI-2がページ本文として資料リンクを回答根拠に拾えるようになる（V15）。C-1・C-2（資料はリンクでありファイル
          実体ではない旨・更新日は任意）はv4までの判断を維持。日本語版／英語版が対の資料は1エントリの説明文に集約。カテゴリに「事例資料」を維持。</div>
        </div>''')(),
        altbox(
            "副案（点線枠・主案と併用可）",
            "副案：リストを併用する案（主案はテキストweb part）",
            "V14の通り、カスタムリストの各アイテムは既定でMicrosoft Search（サイト内検索）にインデックスされるため、<b>人間のサイト内検索・"
            "一覧管理には有効</b>。ただし<b>SharePoint agent（AI-2）の回答根拠にはならない</b>（V15：リストは非対応）。主案（テキストweb part）"
            "を残したまま、必要であれば下記のような<b>リスト web part（カテゴリ／言語／更新日の列＋ビュー絞り込みチップ）を補助的に併設</b>する"
            "運用も可能。<b>主案と副案のどちらか（または両方）を採用するかは社長の判断</b>とする。",
            extra='''<div class="listpart" style="margin-top:12px; background:#fff;">
              <div class="filterbar">
                <span class="lbl">ビューで絞り込み（人間向け・補助）:</span>
                <span class="chip">共通</span><span class="chip">ソリューション</span><span class="chip">アカデミー</span>
                <span class="chip">事例資料</span><span class="chip">お役立ち資料</span>
              </div>
              <table>
                <tr><th>タイトル（ハイパーリンク列）</th><th>カテゴリ（選択肢列）</th><th>言語（選択肢列）</th><th>更新日（任意列）</th></tr>
                <tr><td class="linkcol">会社案内（日本語版）</td><td class="catcell">共通</td><td>日本語</td><td>—</td></tr>
                <tr><td class="linkcol">ソリューション紹介資料（約11本）</td><td class="catcell">ソリューション</td><td>日本語／英語</td><td>—</td></tr>
                <tr><td class="linkcol">お役立ち資料（30本超）</td><td class="catcell">お役立ち資料</td><td>日本語／英語</td><td>任意</td></tr>
              </table>
              <div class="footnote1">※このリストはAI-2の回答根拠にはならない（V15）。人間の検索・一覧管理の補助としてのみ機能する。</div>
            </div>'''
        ),
        sec(
            5, "メルマガ配信スケジュール", "[クイックリンク web part]",
            "現状維持（社外秘表記も維持）。",
        ),
        changes_box([
            "【v5・最重要】「公開資料ライブラリ」を<b>リスト web partからテキスト web part（主案）に変更</b>した（V15：AI-2はリストの"
            "データを回答根拠に使わないと公式確認。カテゴリ見出し＋リンク箇条書きで表現し、お役立ち資料は折りたたみ格納。トレードオフ"
            "（カテゴリ/言語の動的な絞り込み・並べ替えを失う）を明記。<b>副案（点線枠）としてリスト併用案</b>を併記し、社長が主案/副案を"
            "選べる形にした）。",
            "【v4】「AKKODiS企業向けWebページ」（セミナー／活用事例／Insights／サービス）を<b>クイックリンク4列表示に戻した</b>（V8再検証：3列上限はセクション列レイアウトのみに適用）。",
            "（v4までの継続）3列制限はリストの列表示ではなくビュー絞り込みのため元々非該当。旧SECTION5案内は枠内の1行注記に簡素化済み。「事例資料」カテゴリ維持。",
        ]),
        footnote("ToBe WF v5 — 3. クライアント向け — SharePoint UI再検証 v5（V15）反映。旧V8も反映継続"),
    ]
    return render("client", 3, TOTAL_PAGES, "3. クライアント向け", blocks)


# ---------------------------------------------------------------------------
# 4. キャンディデート向け
# ---------------------------------------------------------------------------
def build_candidate():
    blocks = [
        hero("キャンディデート向けページ", "現状維持"),
        toc([
            {"label": "AKKODiS公式HP", "target": "AKKODiS People（テキストweb part・複数ハイパーリンク）／リクルーティングムービー（クイックリンク）"},
            {"label": "マスター資料", "target": "会社説明資料（現行踏襲・3チーム連絡先）"},
            {"label": "Social Media", "target": "X／Facebook／LinkedIn／YouTube（クイックリンク4列）"},
        ]),
        resolved_notice(
            "本ページは v5 でも変更なし（トップ／キャンディデート／パートナーは変更なし）",
            "「AKKODiS People」のテキスト化（クイックリンク→テキストweb part＋複数ハイパーリンク）と「Social Media」の4列表示は、"
            "<b>いずれもv4で対応済み</b>（下記SECTION2・SECTION4に反映済み）。社長より同一の指摘が再送されたため、既に対応済みである"
            "ことが一目で分かるよう本ボックスを追加した。追加の修正は行っていない。",
        ),
        (lambda: f'''<div class="sec">
          <div class="sechead">
            <div>
              <div class="num">SECTION 2</div>
              <div><span class="part-tag">①[テキスト web part]（複数ハイパーリンク）　②[クイックリンク web part]</span></div>
              <div class="title">AKKODiS公式HP（v4：AKKODiS Peopleをテキスト化／v4で対応済み）</div>
            </div>
            <div class="badges"><span class="badge v4">v4改訂</span><span class="badge rule">V11</span><span class="badge ok">v4で対応済み</span></div>
          </div>
          <div class="desc">【v4】旧v3は「AKKODiS People」のクイックリンクのタイルに、業務依頼フォームへのリンクを説明文脈内リンクとして添えていたが、
          V11検証の結果<b>クイックリンクの1タイル＝1リンクが構成上の制約</b>（画像・アイコン・説明文・代替テキストの設定項目のみで、副リンクを追加する機能は無い）と判明。
          「AKKODiS People」はクイックリンクのタイルから<b>テキスト web part</b>に変更し、説明文＋複数のハイパーリンクとして表現する
          （クイックリンクのタイルに副リンクを入れる旧設計の誤りを解消）。</div>
          <div class="cols2">
            <div class="col">
              <div class="t">[テキスト web part] AKKODiS People</div>
              <div class="desc" style="margin-bottom:8px;">AKKODiSで働く「人」やカルチャーを紹介する採用サイト（現状維持）。</div>
              {linkgroups_block([
                  ("本文中の複数ハイパーリンク", ["AKKODiS Peopleを見る（外部サイトへ）", "※告知投稿の依頼は業務依頼フォームから（マーケへの業務依頼フォームへ）"]),
              ], single_col=True)}
            </div>
            <div class="col">
              <div class="t">[クイックリンク web part] リクルーティングムービー</div>
              <div class="card" style="border:none; padding:0;">
                <div class="icon"></div>
                <div class="d">現状維持。1タイル＝1リンクの通常のクイックリンクとして配置する。</div>
              </div>
            </div>
          </div>
          <div class="note">2カラムはWF上の便宜的なレイアウト（SharePointのセクション2列レイアウトでも表現可）。
          「AKKODiS People」＝テキストweb part、「リクルーティングムービー」＝クイックリンクweb partと、<b>パーツの種類が異なる</b>点に注意（V11）。</div>
        </div>''')(),
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
            4, "Social Media（v4で4列対応済み）", "[クイックリンク web part]",
            [("X", "現状維持"), ("Facebook", "現状維持"), ("LinkedIn", "現状維持"), ("YouTube", "現状維持")],
            note="【v4】4列に戻した：V8再検証によりクイックリンクweb part自体はレスポンシブに4個横並び表示が可能と確認"
            "（X／Facebook／LinkedIn／YouTubeの4タイル横並び）。",
            badges=[("v4改訂", "v4"), ("V8", "rule"), ("v4で対応済み", "ok")],
            cols=4,
        ),
        changes_box([
            "【v5】本ページは<b>変更なし</b>（v5変更指示：トップ／キャンディデート／パートナーは変更なし）。"
            "冒頭に「v4で対応済み」の強調ボックスを追加し、社長からの同一指摘（AKKODiS Peopleのテキスト化・Social Media 4列）が"
            "<b>既にv4で反映済み</b>であることを一目で分かるようにした（内容自体の追加変更はなし）。",
            "【v4】「AKKODiS公式HP」の「AKKODiS People」を、クイックリンクのタイルから<b>テキストweb part（説明文＋複数ハイパーリンク）に変更</b>した"
            "（V11検証：クイックリンクは1タイル＝1リンクの構成上の制約があり、業務依頼フォームへの副リンクをタイル内に入れられないため）。",
            "【v4】「Social Media」を<b>クイックリンク4列表示に戻した</b>（V8再検証：3列上限はセクション列レイアウトのみに適用）。",
            "（v3までの継続）CA-1（マスター資料）は現行文言をそのままテキストweb部品に確定掲載する設計を維持。CA-2（3セクション構成）は変更なし。",
        ]),
        footnote("ToBe WF v5 — 4. キャンディデート向け — v5変更指示：本ページは変更なし（AKKODiS Peopleテキスト化・Social Media 4列はv4で対応済み）"),
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
            "【v5】本ページは<b>変更なし</b>（v5変更指示：トップ／キャンディデート／パートナーは変更なし）。",
            "【v4】本ページは変更なし（v4変更指示：パートナーはv3のまま）。",
            "（v3までの継続）3パートナーのサムネイルカード（役割A）を削除し、ページ内目次（クイックリンク web part）のみを残した。P-1（ASP Teamの画像枠是正）・P-2（ページ内目次新設）は維持。",
        ], head="このページの変更点（v5：変更なし・v3から継続）"),
        footnote("ToBe WF v5 — 5. グローバルブランドパートナーシップ（v3から変更なし）"),
    ]
    return render("partner", 5, TOTAL_PAGES, "5. グローバルブランドパートナーシップ", blocks)


# ---------------------------------------------------------------------------
# PDF結合（1ページ＝1WF）
# ---------------------------------------------------------------------------
def build_pdf(png_paths, out_name="ToBe_WF_v5_承認用.pdf"):
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

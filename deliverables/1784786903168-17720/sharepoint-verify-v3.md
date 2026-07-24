# SharePoint Online（モダン コミュニケーション サイト）精密再検証 v3

- 検証日: 2026-07-24
- 目的: 依頼者（SharePoint精通者）から前回WFの誤り指摘を受け、V1〜V7の正式名称・数値をlearn.microsoft.com / support.microsoft.comの一次情報で再確定する
- 前回調査ファイル: `/tmp/claude-0/-home-user-ai-corporation/e24d4341-6e4c-5b20-80f3-9be0c89ca4ef/scratchpad/sharepoint-feasibility.md`（R1〜R9、判断軸=L1/L2）
- 方針: 日本語UI表記を優先して確定し、英語名も併記。裏取りできない項目は断定せず「要確認」と明記。

---

## V1. セクションの最大カラム数

| 項目 | 内容 |
|---|---|
| 事実 | モダンページのセクションに追加できる列レイアウトは **1列 / 2列 / 3列 / 全幅列（コミュニケーションサイトのみ）/ 垂直セクション（ページ右側専用）/ フレキシブル セクション / セクション テンプレート** の組み合わせ。**セクションあたりの最大列数は3列**であり、原文に "To show content side-by-side, **you can add up to three columns** to each section." と明記。**4カラムという選択肢はOOTBには存在しない（不可能）** |
| 正式名称（日） | セクション（Section）／列（Column）／全幅列（Full-width column）／垂直セクション（Vertical section）／フレキシブル セクション（Flexible section） |
| 正式名称（英） | Section / Column / Full-width column / Vertical section / Flexible section |
| 出典URL | https://support.microsoft.com/en-us/office/add-sections-and-columns-on-a-sharepoint-modern-page-a54ecb22-ce73-46db-8d08-95a810ffcbde （英語版・原文引用元）／日本語版: https://support.microsoft.com/ja-JP/SharePoint/pages-in-sharepoint/add-sections-and-columns-on-a-sharepoint-modern-page |
| 確度 | **高**（"up to three columns" の一文で明記。4列不可も同一記事から論理的に確定） |
| 実務注意 | 「4カラムで並べたい」という設計要求はOOTBでは実現不可。3列セクション＋別セクション追加（縦積み）で代替するか、フレキシブル セクション内で列比率を変える（例: 1:2, 2:1等の比率変更は可能だが列の本数自体は最大3）方針にする。前回WFで4カラムを前提にした箇所があれば要修正。 |

---

## V2. 「Highlighted content」web partの正式日本語名称

| 項目 | 内容 |
|---|---|
| 事実 | support.microsoft.com の日本語記事タイトルが **「強調表示されたコンテンツの Web パーツを使用する」** であり、これが公式の日本語UI表記。**「強調されたコンテンツ」でも「ハイライトされたコンテンツ」でもなく「強調表示されたコンテンツ」が正式名称** |
| 正式名称（日） | **強調表示されたコンテンツ** Web パーツ |
| 正式名称（英） | Highlighted content web part |
| 出典URL | https://support.microsoft.com/ja-jp/sharepoint/pages-in-sharepoint/use-the-highlighted-content-web-part |
| 確度 | **高**（日本語版support記事タイトルに直接明記） |
| 実務注意 | **前回調査ファイルとの矛盾を指摘**: `sharepoint-feasibility.md` のR2見出し・R6本文では **「強調されたコンテンツ (Highlighted content)」と表記されており「表示」が抜けている**（誤記）。正式には「**強調表示**されたコンテンツ」が正しい日本語名称。WF・提案書内の表記は全て「強調表示されたコンテンツ」に統一修正が必要。 |

---

## V3. 「List」web partの正式日本語名称とハイパーリンク列

| 項目 | 内容 |
|---|---|
| 事実(a) List web part名称 | 日本語版support記事タイトルは **「リスト Web パーツを使う」**。正式名称は **「リスト Web パーツ」**。ビュー選択・フィルター・並べ替え・グループ化が標準機能（"ユーザーは、リストを表示、フィルター、並べ替え、グループ化する" 等の記述あり） |
| 事実(b) ハイパーリンク列 | **実在する**。列タイプとして **モダン体験では「ハイパーリンク」**（英: Hyperlink）という単独の列タイプが存在し、「Web ページ、グラフィック、またはその他のリソースへのハイパーリンクを格納するために使用」と明記。**クラシック体験では「ハイパーリンクまたは画像」（英: Hyperlink or Picture）** という列タイプ名（画像表示機能も統合）で、モダンとクラシックで名称が異なる点に注意。入力時はURLと説明テキスト（表示名）を入力する仕様 |
| 事実(c) 外部URLリンク集の実現性 | **可**。カスタムリストに「ハイパーリンク」列＋カテゴリ/言語等の選択肢列を追加し、リスト標準の**ビュー（View）**機能でフィルター・並べ替え・グループ化した絞り込み表示を作成、ページ上に**リスト Web パーツ**でそのビューを表示、という構成がOOTBで実現可能 |
| 正式名称（日） | リスト Web パーツ（List web part）／列の種類「ハイパーリンク」（モダン）・「ハイパーリンクまたは画像」（クラシック） |
| 正式名称（英） | List web part / Column type "Hyperlink" (modern) / "Hyperlink or Picture" (classic) |
| 出典URL | リスト web パーツ: https://support.microsoft.com/ja-jp/office/リスト-web-パーツを使う-ef0a1b80-f8b3-443d-b04a-1e76c70b5537 ／列の種類: https://support.microsoft.com/ja-jp/sharepoint/lists/data-and-lists/list-and-library-column-types-and-options （日本語版で「ハイパーリンク型」の説明を確認）／英語版（モダン/クラシックの名称差の確認元）: https://support.microsoft.com/en-us/sharepoint/lists/data-and-lists/list-and-library-column-types-and-options |
| 確度 | **高**（列タイプ名・Web part名とも公式記事に明記。ただしモダン/クラシックでの名称差異は英語版記事からの読み取りであり、日本語UI上での「ハイパーリンク」表記がモダン全テナントで完全に統一されているかは**要確認**〔テナントのUIバージョンにより表記ゆれの可能性〕） |
| 実務注意 | 前回調査（R6）の結論「Lists（ハイパーリンク列）+ビュー+List web partで外部リンク集を構造化管理できる」は**本検証でも支持される**。矛盾なし。列作成時のUIでは「列の追加」→種類選択で「ハイパーリンク」を選ぶ操作フローになる。 |

---

## V4. セクションの折りたたみ機能

| 項目 | 内容 |
|---|---|
| 事実 | モダンページの**セクション単位の折りたたみ**は標準機能として存在。編集モードでセクションを選択→プロパティパネルの **「このセクションを折りたたみ可能にする」** トグルをオンにする。セクションタイトルの表示有無・既定の展開/折りたたみ状態・区切り線・アイコン配置などを合わせて設定可能 |
| 用途（格納用途としての可否） | **可**。コンテンツをセクション単位で畳んで隠し、必要時にクリックで展開させる用途に使える |
| 制約 | ①**SharePoint Online（Microsoft 365）専用機能**。SharePoint Server 2019 / Subscription Editionでは非対応。②粒度は「ページのセクション」という大きな単位に限られ、コンポーネント内の**多数項目を個別にアコーディオン開閉する**用途（FAQリストのような細粒度UI）はこの機能の対象外（前回調査R5と同じ結論、矛盾なし）。③複数セクションはそれぞれ独立してプロパティ設定できるため、ページ内の複数箇所を個別に折りたたみ配置することは可能 |
| 正式名称（日） | **このセクションを折りたたみ可能にする**（セクションのプロパティ内トグル） |
| 正式名称（英） | Make this section collapsible |
| 出典URL | https://support.microsoft.com/en-us/office/add-sections-and-columns-on-a-sharepoint-modern-page-a54ecb22-ce73-46db-8d08-95a810ffcbde （英語原文: "In the Section properties, under **Make this section collapsible**, switch the toggle on."）／日本語版: https://support.microsoft.com/ja-JP/SharePoint/pages-in-sharepoint/add-sections-and-columns-on-a-sharepoint-modern-page |
| 確度 | **高**（トグル名・制約とも公式記事に明記。前回調査R5と結論一致、矛盾なし） |
| 実務注意 | WF上「アコーディオンで畳む」という表現を使う場合は、正式名称「セクションの折りたたみ（このセクションを折りたたみ可能にする）」であり、一般的な「アコーディオンUI（項目ごとの個別開閉）」とは異なる粒度である旨を注記する。多数項目の個別開閉が必須ならSPFxカスタム開発が必要（L2以上、前回調査と同じ）。 |

---

## V5. クイックリンク(Quick Links) web partの見出しテキスト（長文時の挙動）

| 項目 | 内容 |
|---|---|
| 事実 | learn.microsoft.com / support.microsoft.com の一次記事本文には、**見出し（タイトル/リンクテキスト）の文字数上限や長文時の省略・折り返し挙動について明記した記述が見つからなかった**。公式記事はレイアウトの見た目（画像例）と、一部レイアウトのアイコン/画像の**ピクセルサイズ**（例:コンパクト=48×48px、映写フィルム=幅212〜286px）のみを記載し、テキストの文字数制限には言及していない |
| 二次情報（参考・未確定） | 複数のサードパーティ実務ブログ（例: LookBook365, EnjoySharePoint等）では「タイトルは約110文字まで入力可能」「2行を超える説明文は省略記号(…)で切られ全文を見る手段がない」といった実務報告があるが、**これらはMicrosoft公式一次情報ではないため断定不可** |
| 正式名称（日） | クイック リンク Web パーツ |
| 正式名称（英） | Quick Links web part |
| レイアウト種別（公式確認済み） | コンパクト(Compact) / 映写フィルム(Filmstrip) / グリッド(Grid) / ボタン(Button) / リスト(List) / タイル(Tiles) の6種類 |
| 長文リンクに向くパーツ | 公式記事に直接の比較記載はないが、Quick Linksは短い見出し前提の視覚的パーツであるのに対し、**長い見出し（20〜40字級）を安全に収めたい場合はList web part（リスト。列幅が可変でテキスト折り返し前提）またはText web part内のハイパーリンク付き箇条書き/表**の方が適する（前回調査R6の結論とも整合） |
| 出典URL | https://support.microsoft.com/en-us/office/use-the-quick-links-web-part-e1df7561-209d-4362-96d4-469f85ab2a82 ／日本語版: https://support.microsoft.com/ja-jp/office/クイックリンク-web-パーツを使用する-e1df7561-209d-4362-96d4-469f85ab2a82 |
| 確度 | **低〜中（文字数の数値は要確認）**。レイアウト種別・ピクセルサイズの一部は一次情報で確度高。**「20〜40字級の見出しが収まるか」については一次情報で断定できず、要確認のまま明記する**。実機テナントでの実機検証（実際に長い見出しを入れて挙動を見る）を推奨 |
| 実務注意 | 前回調査（R6）は「Quick Linksはカテゴリ絞り込みに不向き」という別論点のみを扱っており、本項目（見出し文字数）は前回未検証・矛盾なし（新規追加項目）。**WFで長い見出しを使う設計をする場合は、Quick Linksを避けてList web partまたはText web partを使う方針が安全**（断定できない以上、リスクを取らない設計を推奨）。 |

---

## V6. テキスト(Text)web partでのハイパーリンク挿入

| 項目 | 内容 |
|---|---|
| 事実 | **可**。Text web part内でハイパーリンクを挿入する方法は2通り公式に明記: ①ツールバーの「ハイパーリンク」コマンド（またはCtrl+K）をクリックし、挿入リンクダイアログでアドレス（URL）と表示テキストを入力する。②URLをコピーしてテキストボックスに直接貼り付けると自動的にハイパーリンク化される |
| 表・リスト形式での整理 | **可**。Text web partは表（テーブル）機能を持ち、セル内にもハイパーリンクを挿入できるため、**表形式でのリンク一覧整理も標準機能で可能**（前回調査R7で確認済みの「テーブルのセル背景色」機能とも整合し、Text web part内でテーブル機能自体が使えることの裏付けにもなる）。箇条書き（番号なし/番号付きリスト）内へのハイパーリンク挿入も同様に可能 |
| 見出しとの連動 | 見出し（Heading 1〜3）にはページアンカー（bookmark）が自動生成される（前回調査R4と同じ結論） |
| B-2方針（既存資料へのリンクをテキストパーツで貼る）の妥当性 | **妥当**。新規ライブラリを作らずとも、Text web part内に既存ドキュメントの共有URLを表・箇条書き形式で整理して貼る運用はOOTBで問題なく実現できる |
| 正式名称（日） | テキスト Web パーツ |
| 正式名称（英） | Text web part |
| 出典URL | https://support.microsoft.com/ja-jp/office/テキスト-web-パーツを使用して、テキストやテーブルをページに追加する-729c0aa1-bc0d-41e3-9cde-c60533f2c801 ／英語版: https://support.microsoft.com/en-us/office/add-text-tables-and-images-to-your-page-with-the-text-web-part-729c0aa1-bc0d-41e3-9cde-c60533f2c801 |
| 確度 | **高**（ハイパーリンク挿入の2手段とも公式記事本文に明記。前回調査と矛盾なし） |

---

## V7. 既存ドキュメントライブラリの参照

| 項目 | 内容 |
|---|---|
| 事実 | **可**。「新規ライブラリを作らず既存ライブラリを参照する」標準手段は主に2つ:<br>①**ドキュメント ライブラリ Web パーツ**をページに追加し、web partの編集設定で**既存の**ドキュメントライブラリを選択（"ドキュメント ライブラリを選択してください" 相当の操作）。表示するビューも事前作成した既存ビューから選択可能。特定フォルダのみを表示する設定（フォルダ名を`/`区切りで指定）も可能。ユーザーは適切な権限があればweb part上から直接ファイルの表示・編集、または「すべて表示」で完全なライブラリへ遷移できる。<br>②**Text web part内にファイルへの直接リンク（URL）を貼る**（V6で確認済み） |
| クロスサイト参照の可否 | 公式記事本文には「別サイトのライブラリを選択できる」旨の明記は確認できず、**同一サイト内のライブラリを対象とする記述が中心**（クロスサイトでのweb part参照可否は**要確認**。実務上は別サイトのライブラリの場合はTextパーツでのリンク方式が確実） |
| 正式名称（日） | ドキュメント ライブラリ Web パーツ |
| 正式名称（英） | Document Library web part |
| 出典URL | https://support.microsoft.com/en-us/office/use-the-document-library-web-part-a9dfecc3-2050-4528-9f00-2c5afc5731b0 ／日本語版（Web パーツ一覧内での言及）: https://support.microsoft.com/ja-jp/sharepoint/web-parts-and-apps-in-sharepoint/using-web-parts-on-sharepoint-pages |
| 確度 | **高**（既存ライブラリ・既存ビュー・特定フォルダの指定が可能な点は公式記事に明記。クロスサイト参照の可否のみ要確認） |
| 実務注意 | 前回調査（R3）の結論「ドキュメントライブラリweb partで既存ライブラリのビューを表示できる」は**本検証でも支持される**（矛盾なし）。同一サイト内なら Document Library web part、別サイトや単純なリンク列挙なら Text web part（またはList web partのハイパーリンク列）という使い分けをWFに明記すべき。 |

---

## 前回調査（sharepoint-feasibility.md）との矛盾点まとめ

| 項目 | 前回の表記 | 本検証の結論 | 対応 |
|---|---|---|---|
| Highlighted content の日本語名称 | 「強調**された**コンテンツ」（R2見出し・R6本文） | 「強調**表示**されたコンテンツ」が公式（support.microsoft.com記事タイトルに明記） | **要修正**。WF・提案書内の表記を全て「強調表示されたコンテンツ」に統一する |
| その他（V1・V3・V4・V6・V7） | 該当項目は前回R3/R5/R6/R7と結論一致、または前回未検証の新規項目（V1・V5） | 矛盾なし／新規確定 | 特になし |

---

## 確度サマリ

| ID | 結論 | 確度 |
|---|---|---|
| V1 セクション最大3列（4列は不可） | 不可（4列） | 高 |
| V2 Highlighted content の日本語正式名称 | 「強調表示されたコンテンツ」 | 高 |
| V3 List web part日本語名称＋ハイパーリンク列の実在 | 「リスト Web パーツ」／実在（可） | 高（モダン/クラシックの名称差異のみ要確認） |
| V4 セクション折りたたみの正式名称・可否 | 「このセクションを折りたたみ可能にする」／可（大枠単位のみ） | 高 |
| V5 Quick Links見出しの文字数・挙動 | **要確認**（一次情報に数値記載なし） | 低〜中 |
| V6 Text web partでのハイパーリンク挿入 | 可（表・箇条書き内も可） | 高 |
| V7 既存ドキュメントライブラリの参照 | 可（同一サイト内）／クロスサイトは要確認 | 高（一部要確認） |

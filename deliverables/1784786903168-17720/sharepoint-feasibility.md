# SharePoint Online（モダン コミュニケーション サイト）実現性調査

- 調査日: 2026-07-24
- 調査目的: 社内マーケティングポータル刷新提案の ToBe 設計における「SharePointの既存パーツ（L1=Out-of-the-box）で実現できるか」の判断材料
- 判断軸: L1（既存パーツ・OOTB設定）で実現できるか／L2以上（PnP等コミュニティ製・SPFxカスタム開発・第三者ツール）が必要か
- 出典は learn.microsoft.com を最優先とし、次点で support.microsoft.com（Microsoft公式エンドユーザー向けヘルプ）を使用。いずれも一次情報。到達・fetch確認済みのものにはURLを明記。

---

## R1. 検索ボックスの設置位置（最重要）

| 項目 | 内容 |
|---|---|
| (a) ヘッダー検索ボックスの既定位置 | **可（既定でヘッダー上部に表示）**。"The most visible difference is that the Microsoft Search box is placed at the top of SharePoint, in the header bar."（モダン検索＝Microsoft Searchの説明として明記） |
| 使用パーツ | サイトヘッダーの組み込み検索ボックス（Microsoft Search／suite navigation bar 検索ボックス）。ページ上に追加する「web part」ではなく、サイトのチュームUIの一部 |
| 根拠URL | https://learn.microsoft.com/en-us/sharepoint/guide-to-sharepoint-modern-experience （"Search" 節） |
| 確度 | 高（公式一次情報で明記） |
| (b) ページ本文（コンテンツ領域）に置ける「検索ボックス web part」の有無 | **不可（標準機能としては存在しない）**。Microsoft Q&A（公式フォーラム、learn.microsoft.com配下）で "Site Page is a modern page that does not support search web parts."（モダンページは検索web partをサポートしない）と明記。学習記事側でも、classic の Web Part Page（classic版）でのみ標準の Search Box Web Part が使える旨が確認できる |
| 使用パーツ | 標準では該当パーツなし。回避策として **PnP Modern Search Web Parts**（Search Box / Search Results 等）が存在するが、これは **Microsoft公式サポート対象外のオープンソース・コミュニティ製ソリューション**（"open-source project built by Microsoft together with MVPs and community members"、"wholly maintained by the enthusiastic SharePoint community"）であり、SPFxパッケージのテナントへのインストール（App Catalog経由の配布）が必要 = L2相当（標準OOTBではない） |
| 根拠URL | https://learn.microsoft.com/en-us/answers/questions/2077462/adding-a-search-function-to-the-page 、https://learn.microsoft.com/en-us/answers/questions/286129/search-webpart-in-modern-sharepoint-page.html （Microsoft Q&A＝公式フォーラム）／PnP Modern Search: https://microsoft-search.github.io/pnp-modern-search/ （非learn.microsoft、参考情報） |
| 確度 | 中〜高（Microsoft Q&A は公式フォーラムだが個別回答であり、Learn記事本文ほどの一次性はやや落ちる。ただし「Search Box Web Part」がClassic専用である点は https://learn.microsoft.com/en-us/sharepoint/search-box-web-part の記述（"Settings"→"Edit Page"→"Web Part Menu" というクラシックUI手順のみが記載）とも整合するため確度は高いと判断） |
| (c) 各ページの任意位置（右上等）に検索窓を置けるか | **不可（標準機能では位置変更不可）**。ヘッダー検索ボックスの表示/非表示は `Set-PnPSearchSettings -Scope Web -SearchBoxInNavBar Hidden` 等のPnP PowerShell（非公式=サードパーティOSSツールだが実務上デファクト）で可能だが、**位置そのものはサイトのチュームUI（ヘッダー）に固定**されており移動不可。同記事に "This setting only applies to the search box in the suite navigation bar. It doesn't apply to search boxes that are in the page" とあり、ページ内に検索ボックスを置く行為自体が標準外であることが裏付けられる |
| 根拠URL | https://learn.microsoft.com/en-us/microsoftsearch/manage-spo-search-box |
| 確度 | 高 |

**結論（依頼者の認識の正否）**: 依頼者の「ヘッダー/右上には検索窓を設置できないはず」という認識は**ほぼ正しい**。正確には「①既定のMicrosoft Search検索ボックスはヘッダー上部に**自動で**表示される（位置固定・移動不可）」「②ページ本文側・右上などの任意位置に検索ボックスをOOTBで新設することはできない（標準web partが存在しない）」の2点が learn.microsoft.com / Microsoft Q&A で裏付けられる。実現するには PnP Modern Search 等の非公式導入かSPFxカスタム開発が必要（L2以上）。

---

## R2. 「強調されたコンテンツ (Highlighted content)」web part の絞り込み

| 項目 | 内容 |
|---|---|
| 実現可否 | **条件付き可** |
| 内容 | コンテンツタイプ自体を直接フィルタ条件にする専用UIはないが、**Source選択時にコンテンツタイプ相当のType指定＋Managed Property（管理プロパティ）によるフィルタ**が可能。ドキュメントライブラリ以外の全Sourceオプションで管理プロパティフィルタが使用可能、と明記。標準フィルタ項目は Title includes the words / Content includes the words / Recently added / Recently changed / Created by / Modified by の6種類＋Managed Property指定、AND/ORの組み合わせが可能 |
| 使用する標準パーツ名 | **強調されたコンテンツ (Highlighted content) web part** |
| 根拠URL | https://support.microsoft.com/en-us/office/use-the-highlighted-content-web-part-e34199b0-ff1a-47fb-8f4d-dbcaed329efd |
| 確度 | 中（要素の存在は一次情報で確認済みだが、「サイトのカスタム列（サイト列）をコンテンツタイプに紐付けてManaged Propertyとして検索クロール・マッピングする」設定作業が必要で、テナント側の検索スキーマ設定に依存する。実運用時は「要テナント確認」） |
| 実務上の注意 | カスタム列を Managed Property として使うには、サイト列作成→クロールドプロパティへのマッピング→検索インデックス再クロール、という手順が必要（即時反映されない場合がある）。テンプレートをコンテンツタイプで絞り込み表示、という設計は技術的に実現可能だが、**検索インデックスの反映ラグ**が運用上のリスクとして残る |

---

## R3. ドキュメントライブラリ web part＋メタデータ列

| 項目 | 内容 |
|---|---|
| 実現可否 | **条件付き可** |
| カスタム列（対象/言語/種別/版/更新日）の付与 | 可。SharePointリストの標準機能（列追加は Lists / Library の基本機能。ハイパーリンク列を含む各種列タイプが標準サポート） |
| ライブラリ本体でのビュー/フィルター/列絞り込み | 可。ライブラリ本体（フルUI）ではビュー作成・列フィルター・グループ化が標準機能 |
| ページ上の Document Library web part での絞り込み | **条件付き可**。web part上でも「事前に作成したビューを選択して表示」は可能（"you can choose to show a specific view of the library"）。また列見出しの矢印から**その場での sort / filter / group** は可能。ただし **web part上での操作はセッション限りで元のライブラリには影響しない**（"actions taken in the web part do not affect the underlying library"）＝ページ訪問者ごとの一時的な絞り込みであり、恒久的な「絞り込み済みビュー」をURLで固定共有する用途には別途ビュー作成が必要 |
| 使用する標準パーツ名 | **ドキュメント ライブラリ (Document Library) web part** |
| 根拠URL | https://support.microsoft.com/en-us/office/use-the-document-library-web-part-a9dfecc3-2050-4528-9f00-2c5afc5731b0 、列種別: https://support.microsoft.com/en-us/sharepoint/lists/data-and-lists/list-and-library-column-types-and-options |
| 確度 | 高（web partの基本仕様として一次情報で確認）。ただし「web part上の複数条件同時フィルター」の制約（別調査ではLibrary/List web partの動的フィルタリングは既定で1条件のみ、という報告あり）は**要テナント確認**（Microsoft Q&Aベースの情報であり確定情報ではない） |
| 実務上の注意 | 「対象/言語/種別/版/更新日」ごとに複数ビューを事前設計し、ページ上に複数のDocument Library web part（ビュー違い）を配置する、または見出しごとにセクションを分ける設計が現実的。web part単体でのエンドユーザー向け動的多条件フィルタUIはOOTBでは弱い |

---

## R4. ページ内アンカー／目次（最重要）

| 項目 | 内容 |
|---|---|
| 実現可否 | **可** |
| セクションへのアンカーリンク | 可。**Text web part内の見出し（Heading 1〜3）に対して、SharePointがページアンカー（bookmark）を自動生成**する。"Page anchors (also known as bookmarks) are automatically added to Heading 1, Heading 2, and Heading 3 styles in Text web parts on your page."。公開後のページで見出しにホバーするとリンクアイコンが表示され、右クリック→リンクをコピーで取得できる |
| ネイティブの目次機能 | **専用の自動目次(TOC)web partはOOTBには存在しない**。目次は手動で作る必要がある |
| Quick Links web partから同一ページ内セクションへ飛べるか | **可（ただし公式ドキュメントに明記の専用機能ではなく、汎用リンク機能の応用）**。Quick Links web partの追加リンク先種別は「最近使用したファイル／サイト／OneDrive／コンピューター／リンク（URL）」であり、「リンク」としてページ内アンカーURL（コピーしたページアンカーのURL）を貼り付けることで実現可能。ただしMicrosoft公式記事に「同一ページ内アンカーへのリンク」という用途が名指しで書かれているわけではなく、一般的な「リンク」機能の応用としての実現である点に留意（＝実装は可能だがドキュメントに直接の明記なし＝推測を含む） |
| 使用する標準パーツ名 | **Text web part（見出しでアンカー自動生成）＋ Quick Links web part（アンカーURLへのリンク集＝目次代わり）** |
| 根拠URL | https://support.microsoft.com/en-us/office/add-text-tables-and-images-to-your-page-with-the-text-web-part-729c0aa1-bc0d-41e3-9cde-c60533f2c801 （アンカー自動生成）／ https://support.microsoft.com/en-us/office/use-the-quick-links-web-part-e1df7561-209d-4362-96d4-469f85ab2a82 （Quick Links機能） |
| 確度 | 高（アンカー自動生成は明記）／中（Quick Linksでのアンカー活用は組み合わせとして技術的に成立するが、公式記事内に「ページ内目次としての利用」という直接の記述はなく実運用での見え方は要検証） |
| 実務上の注意 | アンカーは **Text web part内の見出しのみ**に生成され、他のweb part（Highlighted content, Document Libraryなど）には生成されない。ToBeでセクション先頭に必ずText web part＋見出しを置く設計にする必要がある。「目次アンカー」をWFに描く技術的根拠として成立する |

---

## R5. 折りたたみ（アコーディオン/コラプシブル）セクション

| 項目 | 内容 |
|---|---|
| 実現可否 | **条件付き可** |
| 内容 | モダンページの**セクション単位の折りたたみはネイティブ対応**（"Make this section collapsible" トグルをセクションのプロパティで有効化）。タイトルの有無、初期表示状態（展開/折りたたみ）、アイコンの左右配置、区切り線表示などを設定可能 |
| 制約 | ①SharePoint Server 2019 / Subscription Edition では非対応（**SharePoint Online専用機能**）。②**複数セクションを同時に独立して開閉する「アコーディオン」的な一覧UI**（例: FAQリストのように多数の項目を1つのコンポーネント内で個別開閉）はOOTBの「セクション折りたたみ」機能の対象外（あくまで「ページのセクション」という大きな単位の開閉であり、コンポーネント内の複数アイテムの個別アコーディオンではない）。③Document Library / Quick Links / Site Activity 等のコンポーネント単位では折りたたみ設定不可（セクションのみ対応） |
| 使用する標準パーツ名 | ページの**セクション (Section) プロパティ**（web partではなくページレイアウト機能） |
| 根拠URL | https://support.microsoft.com/en-us/sharepoint/pages-in-sharepoint/add-sections-and-columns-on-a-sharepoint-modern-page |
| 確度 | 高（手順・制約とも公式記事に明記） |
| 実務上の注意 | B-7で想定する「アコーディオンで多数のFAQ項目を個別開閉」のような細粒度UIは、セクション折りたたみでは代替しきれない（セクション単位が粒度の限界）。多数項目の個別開閉が必須要件ならSPFxカスタム開発が必要（L2以上）。「大きなブロック単位の開閉」であればOOTBのセクション折りたたみで十分 |

---

## R6. 外部リンク集の整理

| 標準手段 | 絞り込み可否 | 詳細 | 根拠URL |
|---|---|---|---|
| **Quick Links web part** | **不可（カテゴリ分けの絞り込み表示機能なし）** | リンクを見た目良く並べる複数レイアウト（Compact/Filmstrip/Grid/Button/List/Tiles）は提供されるが、**カテゴリタグによる動的フィルタ・絞り込みUIは公式記事に記載なし**。カテゴリ分けは「セクションを分けて複数のQuick Links web partを配置する」という静的な構造化でのみ対応可能（推測を含む） | https://support.microsoft.com/en-us/office/use-the-quick-links-web-part-e1df7561-209d-4362-96d4-469f85ab2a82 |
| **Lists（リスト）のハイパーリンク列＋ビュー** | **可** | Lists標準の「ハイパーリンク列（Hyperlink column）」でURL＋表示テキストを保持でき、加えてカテゴリ/言語/種別等の追加列（選択肢列等）を付与し、**ビュー（View）の並べ替え・グループ化・フィルターの標準機能**で絞り込み表示が可能。ページには **List web part** で配置できる | https://support.microsoft.com/en-us/sharepoint/lists/data-and-lists/list-and-library-column-types-and-options 、https://support.microsoft.com/en-us/office/use-the-list-web-part-ef0a1b80-f8b3-443d-b04a-1e76c70b5537 、ビュー: https://support.microsoft.com/en-us/sharepoint/lists/data-and-lists/create-change-or-delete-a-view-of-a-list-or-library |
| **強調されたコンテンツ web part（Source: リンク相当）** | 参考: リンク集自体を直接ソースにする設計は非標準 | Highlighted content web partはドキュメント/ページ/ニュース等のSharePointコンテンツを対象とし、**外部URLの集合を管理・フィルタする用途には不向き**（Listsの方が適する） | 上記R2出典と同じ |

**結論**: 外部リンク集をカテゴリ/メタデータで整理・絞り込みたい場合は **Lists（カスタムリスト）のハイパーリンク列＋ビュー＋List web part** の組み合わせがOOTBで最も確実。Quick Links web partは見た目は良いが動的な絞り込みには向かない（静的な複数配置での代替のみ）。

---

## R7. テキストのハイライト色（イエローマーカー）

| 項目 | 内容 |
|---|---|
| 実現可否（色の自由指定） | **不可（固定パレットのみ）** |
| 内容 | Text web partの蛍光ペン（ハイライト）機能は**固定の既定パレットのみ**で、カスタムカラーコード指定・パステル/半透明調整には対応しない。モダンリッチテキストエディタは "font, font size, font color, highlight, bold, italics, underline, bullets, numbering" 等の基本書式のみをサポート |
| 使用する標準パーツ名 | **Text web part**（ハイライト機能） |
| 根拠URL | https://support.microsoft.com/en-us/office/add-text-tables-and-images-to-your-page-with-the-text-web-part-729c0aa1-bc0d-41e3-9cde-c60533f2c801 |
| 確度 | 中（固定パレットである点はMicrosoft Q&A等の複数の実務報告と整合するが、learn.microsoft.com/support.microsoft.com本文に「パレット数・色の一覧」が明記されたページは今回確認できず。**具体的な色数・色見本は要テナント確認**） |
| 標準の代替手段 | ①**テーブルの単一セルに背景色を設定**（Text web partのテーブル機能＋セル背景色）。②**セクションの背景色設定**（ページのセクションプロパティで背景色を選択可能。ただし選択肢はテーマ由来の一定パレットに準ずる）。③**見出し/太字/斜体などの書式**で視覚的に強調。④注意喚起の定型パーツとしては、SharePoint標準に「コールアウト/メッセージバー」専用web partは存在せず、**Hero web part**（画像＋テキストで強調表示）や**Call to action web part**（ボタン付き訴求）、**タイトル上のテキスト(Kicker)機能**で代替する運用が現実的（いずれも「注意喚起バナー」そのものではなく目的外流用） |
| 追加根拠URL | Hero web part: https://support.microsoft.com/en-us/office/use-the-hero-web-part-d57f449b-19a0-4b0d-8ce3-be5866430645 ／Call to action web part: https://support.microsoft.com/en-us/office/use-the-call-to-action-web-part-e9917310-7543-4fc4-8f3f-d78e46074c00 |
| 実務上の注意 | 「イエローマーカーで強く注意喚起」という体験は、色の自由指定という点では制約があるものの、固定パレット内に黄色系ハイライトが含まれる可能性は高い（要テナント確認）。ブランドカラーに合わせた自由な色指定が必須要件なら、テーブルセル背景色 or セクション背景色の方が制御しやすい |

---

## R8. 端末別コンテンツ出し分け

| 項目 | 内容 |
|---|---|
| 実現可否 | **不可（モダン標準機能では不可。レスポンシブのみ）** |
| 内容 | モダンSharePoint（コミュニケーションサイト含む）は「レスポンシブデザインにより同一コンテンツが画面サイズに応じて見た目を最適化」される設計であり、**PC/スマホで異なるコンテンツ自体を出し分ける機能（Device Channels相当）はモダン体験には存在しない**。Device Channels（デバイスチャネル）機能は**クラシック体験専用**であり、モダン体験では利用不可 |
| 根拠URL | https://learn.microsoft.com/en-us/sharepoint/guide-to-sharepoint-modern-experience （"modern experience lets anyone make attractive, lively sites and pages that are ready for mobile devices" ＝レスポンシブ前提の記述）／クラシック限定の裏付け: https://learn.microsoft.com/en-us/archive/technet-wiki/23157.sharepoint-2013-device-channels 、https://learn.microsoft.com/en-us/sharepoint/administration/plan-for-mobile-views |
| 確度 | 中〜高（「モダン=レスポンシブ」は公式記事に明記。「Device Channelsがクラシック専用でモダンでは使えない」という点は複数のMicrosoft Q&A回答で一致しているが、learn.microsoft.com本文に「モダンでは device channelsが使えない」と直接明記した一次ソースは今回のfetch範囲では確認できず＝**要テナント確認**として明記） |
| 実務上の注意 | 「PC版とスマホ版で見せる情報を変える」設計は、OOTBでは実現不可。同一コンテンツのレイアウト最適化（列の折り返し等）のみ。どうしても出し分けが必要なら、SPFxでのユーザーエージェント判定カスタム実装（L2以上）が必要 |

---

## R9. 利用状況分析（GA無し前提のKPI材料）

SharePoint標準で「サイト利用状況（Site usage）」と「検索利用状況（Search usage reports）」の2系統のレポートが提供される。

### (A) サイト利用状況（Site usage）— ページ/コンテンツ側

| 指標 | 内容 |
|---|---|
| Unique viewers（一意訪問者数） | サイトへの個別訪問者の合計数 |
| Site visits（サイト訪問数） | 同一ユーザーによる同一アイテムへの連続操作は重複除去するアルゴリズムで集計 |
| Average time spent per user | ユーザーあたり平均滞在時間 |
| Popular content（人気コンテンツ） | 閲覧数の多いファイル・ページの一覧（訪問数/一意訪問者数でソート可） |
| Site traffic | サイトトラフィックの推移 |
| Popular Platforms | 利用デバイス/プラットフォームの傾向 |
| Shared with external users | 外部共有状況 |
| ページ/ニュース個別の閲覧数・一意閲覧者数 | 個別ページ（.aspx）単位でのViews/Unique viewersが確認可能 |

参照方法: サイトの「設定」歯車 → **Site usage**。サイト管理者・所有者・メンバー・訪問者が閲覧可能。

根拠URL: https://support.microsoft.com/en-us/office/view-usage-data-for-your-sharepoint-site-2fa8ddc2-c4b3-4268-8d26-a772dc55779e 、https://support.microsoft.com/en-us/office/view-usage-data-for-sharepoint-pages-and-news-e3186199-ccc8-4445-9162-bb1bcec8b7ee

確度: 高（公式サポート記事に指標名が列挙されている）。※データ保持期間・粒度の詳細（例: 過去何日分か）は今回のfetch範囲では確認できず＝**要テナント確認**。

### (B) 検索利用状況（Search usage reports）— モダンサイトの Microsoft Search

| レポート名 | 内容 |
|---|---|
| Query Volume | 検索クエリの実行数の推移。検索活動の多寡の傾向把握に使用 |
| Top Queries | 人気検索クエリ（3回以上検索され、かつ結果クリックがあったクエリが対象） |
| Abandoned Queries | クリック率が低い人気クエリ＝「検索したが選ばれなかった」クエリ一覧 |
| **No Results Queries（ゼロ件検索ログ）** | 検索結果が0件だった人気クエリ一覧。コンテンツ充実化の材料に使用 |
| Impression distribution | 結果タイプ別のインプレッション推移 |

参照方法: サイト設定 → Site collection administration → Microsoft Search → Configure search settings → **Insights**。過去31日（日次）または過去12か月（月次）のデータを閲覧・Excelダウンロード可能。

根拠URL: https://learn.microsoft.com/en-us/sharepoint/view-search-usage-reports-modern-sites

確度: 高（レポート名・定義が表形式で一次情報に明記）。

**結論**: GAが使えない前提でも、**「ページ別閲覧数」「一意閲覧者数」「人気コンテンツ」（Site usage）** と **「検索クエリ数」「人気検索語」「ゼロ件検索クエリ」「検索結果のクリック率(Abandoned)」（Search usage reports）** の2系統をKPI設計の素材として使える。「検索してもヒットしない情報ニーズ」の可視化（ゼロ件検索ログ）はGA代替として特に有用。

---

## ToBe設計への含意（要約）

1. **検索は「ヘッダー検索に一本化」を前提にWFを組む**: モダン標準では検索ボックスをページ本文の任意位置（右上等）にOOTBで新設できない。依頼者の懸念は妥当。「検索窓を右上に置く」設計は避け、既定のヘッダー検索（自動表示・位置固定）を主動線とし、必要ならプレースホルダー文言のカスタマイズ（`SearchBoxPlaceholderText`）程度に留める。ページ内検索が必須要件ならPnP Modern Search導入かSPFxカスタム開発の是非を別途経営判断として提起すべき。
2. **目次アンカーはWFに描いてよい（技術的根拠あり）**: 各セクション先頭にText web part＋見出し(H1〜H3)を必ず配置する設計にすれば、SharePointが自動でページアンカーを生成する。目次はQuick Links web partでアンカーURLを束ねる形で表現できる（「目次web part」という専用パーツではない点は注記が必要）。
3. **絞り込み表示はテンプレート単位で複数ビューを事前設計する前提にする**: Highlighted content／Document Library web partとも「その場での多条件動的フィルタUI」はOOTBでは弱い。コンテンツタイプ・カテゴリごとに事前作成したビュー／web partをセクション単位で並べる設計（＝運用でカバー）を基本方針にする。
4. **アコーディオン(B-7)は「セクション折りたたみ」で代替可能だが粒度に注意**: 大ブロック単位の開閉はOOTB対応。FAQのような多数項目の個別開閉が必須ならカスタム開発の予算判断が必要になる旨をToBe設計書に明記する。
5. **外部リンク集はQuick LinksではなくLists（ハイパーリンク列＋ビュー）で設計する**: カテゴリ・メタデータでの絞り込みが要件に含まれるなら、見た目重視のQuick Linksより構造化データとして扱えるカスタムリストを軸にする。
6. **注意喚起表現はハイライト色の自由指定に依存しない設計にする**: 色は固定パレット。ブランドカラーでの強調が必要な箇所はテーブルセル背景色 or セクション背景色 or Hero/Call to action web partを使う前提でトンマナ設計する。
7. **端末別の出し分けは要件から外す**: モダンはレスポンシブのみでコンテンツ自体の出し分けは不可。PC/スマホで見せる情報を変えたいニーズがあれば、「同一コンテンツをレスポンシブ最適化する」方針に要件を寄せるか、別途カスタム開発の是非を検討する。
8. **KPIはSite usage＋Search usage reportsの2系統を軸に設計する**: 特に「ゼロ件検索クエリ」はGA代替として、コンテンツの過不足を定量的に示せる数少ない標準指標であり、ポータル刷新の効果測定KPIの柱に据えられる。

---

## 全体の確度サマリ（一次情報での裏取り状況）

| ID | 可否 | 確度 |
|---|---|---|
| R1(a) ヘッダー検索の既定表示 | 可 | 高 |
| R1(b) 本文への検索web part設置 | 不可（標準では） | 中〜高 |
| R1(c) 検索窓の任意位置配置 | 不可 | 高 |
| R2 強調されたコンテンツの絞り込み | 条件付き可 | 中 |
| R3 ドキュメントライブラリ+メタデータ | 条件付き可 | 高（一部の制約は要テナント確認） |
| R4 ページ内アンカー | 可 | 高 |
| R4 Quick Linksでのアンカー活用 | 可（応用） | 中 |
| R5 セクション折りたたみ | 条件付き可 | 高 |
| R6 外部リンク集の整理 | Lists+ビューで可／Quick Linksは不可 | 高 |
| R7 ハイライト色の自由指定 | 不可（固定パレット） | 中 |
| R8 端末別コンテンツ出し分け | 不可 | 中〜高（要テナント確認） |
| R9 利用状況分析 | 可（Site usage + Search usage reports） | 高 |

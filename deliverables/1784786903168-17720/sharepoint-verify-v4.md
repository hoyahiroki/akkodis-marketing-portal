# SharePoint Online UI仕様 一次情報ファクトチェック（V8〜V13）

調査日: 2026-07-24 / 対象: SharePoint Online（モダン コミュニケーション サイト）
出典: learn.microsoft.com / support.microsoft.com のみ（一次情報）

---

## V8. クイックリンク(Quick Links) web partのレイアウトとカラム数

**結論**: クイックリンク web part は **6種類のレイアウト**（Compact / Filmstrip / Grid / Button / List / Tiles）を持ち、うち Grid/Button/List/Tiles は「SharePoint only」（オンプレミス版では非対応）。「クイックリンク web part自体が内部で4つ以上横並び可能」という理解は**方向性としては正しい**が、公式ドキュメントは**具体的な列数・ブレークポイント（何個で折り返すか）を数値で明記していない**。Grid layoutは「レスポンシブな画像リンクの壁（responsive wall of image links）」「行と列で上から下へリフローし、少数〜多数のアイテムを収容できる」と説明されるのみで、固定列数の記載なし。ページの「セクション内カラム数（最大3列）」とは別レイヤーの話である点は正しい（セクションはページのレイアウト用グリッド、Quick Links内部の並びはweb part自身のレスポンシブレイアウトで、セクション幅・画面幅・表示設定に応じて自動的に折り返す）。

**使用パーツ**: Quick Links web part（レイアウト: Compact / Filmstrip / Grid / Button / List / Tiles）

**出典URL**:
- https://support.microsoft.com/en-us/office/use-the-quick-links-web-part-e1df7561-209d-4362-96d4-469f85ab2a82（レイアウト一覧・オプション）
- https://support.microsoft.com/en-us/sharepoint/sites-pages/use-the-quick-links-web-part（同内容の別URL）
- https://learn.microsoft.com/en-us/answers/questions/2115804/sharepoint-2019-quick-links-all-layout-options-not（Grid/Button/List/TilesがSharePoint Onlineのみである旨、Microsoft Q&A＝公式ではないが公式挙動の裏付け）
- https://support.microsoft.com/en-us/office/add-sections-and-columns-on-a-sharepoint-modern-page-a54ecb22-ce73-46db-8d08-95a810ffcbde（セクション最大3列の根拠）

**確度**: 中（レイアウト種類・SharePoint Only区分は確度高。「4つ以上横並び」の具体的列数は公式一次情報に数値の明記なし＝**要確認**）

**注意点**: 実務上はGrid/Tiles/Button layoutをフルワイド（1列）セクションに配置すれば画面幅に応じて4〜6個程度が横並びになることが一般的だが、これは公式ドキュメントで保証された仕様値ではなく、レスポンシブ挙動として観測される結果。厳密な列数を握る必要がある依頼（例:「必ず4列で」）には**プレビュー環境での実測**を推奨する旨を依頼者に伝えるべき。

---

## V9. セクション/サブブロックの背景色

**結論**: 背景色（Section background shading）は**セクション単位でのみ設定可能**。1つのテキストweb part内の複数サブブロック（段落・見出しなど）に個別の背景色を付ける機能は、公式ドキュメントに**記載がなく、UI上も存在しない**（テキストweb partの書式設定はフォント・スタイル・箇条書き・リンク・ハイライト等が中心で、ブロック単位の背景色コントロールは提供されていない）。テキストweb part内のテーブルについては、**セルの個別背景色を自由に指定する機能は公式ドキュメントに記載なし**。用意されているのは5種類の**テーブルスタイル**（Plain / Subtle header / Header / Alternating rows / Column header）で、いずれもテーマカラーに連動した定型スタイルであり、任意色でのセル単位塗りつぶしはサポート対象として明記されていない。

**使用パーツ**: セクション（Section background shading）／テキスト web part（テーブルスタイル機能）

**代替案（確認結果に基づく）**:
1. 各サブブロックを**別セクションに分割**し、セクションごとに背景色を設定する（公式にサポートされた方法）
2. テキストweb part内のテーブルは**定型テーブルスタイル**（Header／Theme-colored header等）でコントラストを演出する（セル任意色の直接指定は不可）
3. Call to action web part、または色のある画像/バナーで疑似的に「ブロックごとの色分け」を表現する（公式ドキュメントで直接の代替案として明記されているわけではなく、実務上の代替アイデアという位置づけ＝**要確認**）

**出典URL**:
- https://support.microsoft.com/en-us/office/add-sections-and-columns-on-a-sharepoint-modern-page-a54ecb22-ce73-46db-8d08-95a810ffcbde（セクション背景色＝セクション単位、最大3列）
- https://support.microsoft.com/en-us/office/add-text-tables-and-images-to-your-page-with-the-text-web-part-729c0aa1-bc0d-41e3-9cde-c60533f2c801（テキストweb part・テーブルスタイル5種）
- https://support.microsoft.com/en-us/office/make-your-sharepoint-site-accessible-to-people-with-disabilities-53707eb5-b7b8-4ee0-ae82-9d4d916f7fe1（Section background shadingのコントラスト要件）

**確度**: 高（セクション単位である点、テーブルは定型スタイルのみである点は複数の一次情報で一致）。ただし「セル単位で任意色を指定する機能が完全に存在しない」という**否定の証明**は一次情報に明示の禁止文言があるわけではなく、「機能として案内されていない」ことからの推定＝中〜要確認。

**注意点**: 依頼者がSharePoint精通者とのことなので、「セクション単位」という制約と「テーブルは定型スタイルのみ」という2点は誤解されやすいポイントとして強調すべき。

---

## V10. 強調表示されたコンテンツ(Highlighted content) web partの使い方（複数カテゴリの見せ方）

**結論**: Highlighted content web partは**1つのweb partにつき1つの動的クエリ（ソース＋フィルタ＋並び替え）**で結果一覧を表示する仕組みであり、**結果をカテゴリ別に内部でグループ化・タブ分けして表示する機能はドキュメント上に存在しない**（Sortオプションは Most recent / Most viewed / Trending / Managed property昇順・降順のみで、「Group by」相当の機能は確認できず）。したがって、5カテゴリ（PowerPoint/Word/メール署名/Forms/バーチャル背景）を明確に分けて見せたい場合の標準的な方法は **(b) カテゴリごとに5つのweb partを個別配置し、それぞれに別々のフィルタ条件（コンテンツタイプ or マネージドプロパティ）を設定する**方が確実であり、(a)「1つのweb partにまとめてカテゴリ混在で一覧表示する」ことは可能だが、複数カテゴリを視覚的に分離する目的には不向き（フィルタは同一種類の条件同士がOR、異なる種類の条件同士はANDで結合される仕様のため、"カテゴリA or カテゴリB"は1web part内でも技術的には組めるが、"カテゴリごとに見出しを分けて表示"はweb part単体では不可）。

**使用パーツ**: Highlighted content web part（複数個、ページ内に配置）

**グループ化・フィルタ機能**: フィルタは「タイトル」「コンテンツの単語」「作成者/変更者」「作成日時/更新日時の範囲」「マネージドプロパティ（Document library以外の全ソースで使用可）」。複数フィルタ適用時は「同種フィルタ同士=OR、異種フィルタ同士=AND」。コンテンツタイプは「+ Add content type」で複数追加可能。**グループ表示（Group by）機能は無い**。

**既存ライブラリのファイルをコンテンツタイプ/メタデータで出し分ける現実的な方法**: ①ソースを「A document library on this site」等に指定 → ②コンテンツタイプでフィルタ（複数選択可）、または③マネージドプロパティ（列の値）でフィルタ。カテゴリ列（選択肢列など）を各ファイルに設定済みであれば、web partごとに異なる列値でフィルタをかけて5個配置するのが現実的。

**出典URL**:
- https://support.microsoft.com/en-us/office/use-the-highlighted-content-web-part-e34199b0-ff1a-47fb-8f4d-dbcaed329efd（Source種別、フィルタ種別とAND/OR仕様、Sortオプション、Layoutオプション、Content type追加）

**確度**: 中〜高（フィルタ仕様・Sort仕様は一次情報に明記され確度高。「グループ化機能が存在しない」は同ページ内に**当該機能の記載が無いことからの推定**であり、将来的な機能追加や見落としの可能性はゼロではない＝**要確認**の余地を残す）

**注意点**: 「カテゴリごとに5つのweb part」を推奨する場合、各web partに見出し（テキストweb partやセクション見出し）を添えて視覚的にカテゴリ名を明示する運用が必要（web part自体にはカテゴリ見出しを自動生成する機能はない）。

---

## V11. クイックリンクのタイル構造（副リンクの可否）

**結論**: 公式ドキュメントには「1タイル＝1リンク」と明記した直接の文言はないが、ドキュメント全体の記述（個別リンクのオプションは「画像の変更」「アイコンの書式設定」「説明文の追加」「代替テキストの追加」のみで、各リンク項目に対して**1つのURL・1つの見出し・任意の説明文（1つ）**を設定する構成として説明されている）から、**1タイル内に副リンク（別途のもう1つのハイパーリンク）を追加する機能は無い**という理解で**実務上ほぼ正しい**。ただし、これを明示的に禁止・限定する一次情報の断定文は見つからなかった。

**説明文＋複数リンクを持たせたい場合の正解パーツ**: 依頼の想定通り、**テキスト web part**を使い、説明文の中に複数のハイパーリンク（テキストリンク）を埋め込む方法が最も柔軟で確実（テキストweb partはリッチテキストとして複数のハイパーリンクを自由に配置できる）。他の代替として、Quick Linksのタイルを「AKKODiS People」用と「業務依頼フォーム」用の**2つの独立したリンク項目として並べる**（副リンクではなく同格の別タイルにする）方法もある。

**使用パーツ**: Quick Links web part（1タイル=1リンクの制約）／代替: Text web部（説明文＋複数ハイパーリンク）

**出典URL**:
- https://support.microsoft.com/en-us/office/use-the-quick-links-web-part-e1df7561-209d-4362-96d4-469f85ab2a82（個別リンクの設定項目一覧＝画像・アイコン・説明文・代替テキストのみ）
- https://support.microsoft.com/en-us/office/add-text-tables-and-images-to-your-page-with-the-text-web-part-729c0aa1-bc0d-41e3-9cde-c60533f2c801（テキストweb partでの複数ハイパーリンク埋め込み）

**確度**: 中（機能一覧からの消去法的推定であり、「副リンクを追加できない」と明言する一次情報の直接引用は**見つからず＝要確認**。実務経験・UI構造から見て誤りである可能性は低いが、依頼者がファクトチェックに厳しいとのことなので「明示的な禁止文言は無く、機能一覧に無いことからの推定」と正直に断ることを推奨）

---

## V12. AI-2（SharePoint agent）のリンク探索能力

**結論**:
(a) **生成AI（LLM）を使う**という理解は**正しい**。公式ドキュメントに "Copilot in SharePoint uses AI services to generate responses." と明記されている。SharePoint agentはCopilotと同じ基盤（AI-poweredなアシスタント）で動作する。

(b) 「ページ本文に載っている外部サイトへのハイパーリンクを、ユーザーの質問に対して探し当てて提示できるか」については、**公式ドキュメントに直接の明記が見つからず「要確認」**。確認できた事実は、エージェントは「エージェントのソースに含まれ、かつ利用者が権限を持つ **サイト・ページ・ドキュメント**」及び「Microsoft Graphの最近のファイル操作・プロフィール情報」を回答の根拠にする、という点まで。ページ本文中のリンク（アンカーテキストやURL文字列）を独立した検索対象として扱い、それを回答に含める・提示するという専用の仕様文言は見当たらなかった。一般的なCopilot/RAG系エージェントの動作原理（ページ本文をテキストとしてインデックス→本文中のURL文字列もテキストの一部として取得され得る）から「本文にリンクとして書かれていれば、質問に応じて回答文中で言及・引用され得る可能性は高い」と推測できるが、これは**公式の断定ではなく推測**である。

(c) エージェントが根拠にするもの: **サイト・ページ・ドキュメント（Word/Excel/PDF等）**、および権限に基づくアクセス制御。SharePoint Embedded文脈では「セマンティックインデックス」がRAG（Retrieval Augmented Generation）のグラウンディングに使われるとの記載あり。

(d) 「外部URL自体を回答として返せるか／出典としてページを案内する形か」も、確認できた一次情報の範囲では**明記が見つからず要確認**。SharePoint Embedded系の関連ドキュメントでは「引用URLはナレッジソースの種類によって異なる（例: PDFなら該当ページへのページ内引用）」との記載はあるが、これはSharePoint Embedded（開発者向け文脈）のものであり、通常のSharePointサイトのagentにそのまま同一の挙動が適用されると断定する一次情報は未確認。

**使用パーツ**: SharePoint agent（サイトに自動組み込みのエージェント。Copilotライセンス or 従量課金で利用）

**出典URL**:
- https://support.microsoft.com/en-us/office/frequently-asked-questions-about-copilot-in-sharepoint-eb1b7668-3d98-4a93-98ef-f0c6dfc694f0（"Copilot in SharePoint uses AI services to generate responses."／根拠となる情報源の記述）
- https://learn.microsoft.com/en-us/sharepoint/get-started-sharepoint-agents（利用要件・ライセンス）
- https://support.microsoft.com/en-us/office/get-started-with-agents-in-sharepoint-69e2faf9-2c1e-4baa-8305-23e625021bcf（エージェントの基本動作）
- https://support.microsoft.com/en-us/topic/knowledge-agent-an-overview-c0b1efc3-81d0-4981-8be9-7ba3a75fae15（AI-Ready Content等の概要、引用仕組みの詳細記載はなし）
- （参考・SharePoint Embedded文脈＝完全一致の保証なし）https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/declarative-agent/spe-da-adv（セマンティックインデックス・引用URLの挙動）

**確度**: (a)高 ／ (b)(d)要確認（一次情報に断定的記載なし。推測を交えて回答している旨を明示） ／ (c)中〜高

**注意点**: 依頼者がファクトチェックに厳しいとのことなので、(b)(d)は「一次情報に明記なし＝現状では断定できない」という立場を正直に伝えるべき。実機検証（実際にAKKODiSのSharePointサイトでagentに質問し、外部リンクを含む回答が返るか試す）を推奨する。

---

## V13（KPI補足）. 現行SharePointの検索ログ

**結論**: 現行のSharePoint標準検索（モダンサイトの「このサイトを検索」）でも、**Search usage reports（検索クエリ量・人気クエリ・離脱クエリ・No Results Queries＝ゼロ件検索を含む）は標準機能として記録・閲覧可能**。したがって、刷新前のベースラインとして**既存データがそのまま使える**。

**確認できた具体仕様**:
- 確認場所: **サイト設定 > サイト コレクションの管理 > Microsoft Search > 検索設定の構成 > インサイト（Insights）**
- レポート種別: Query Volume（クエリ量）／Top Queries（人気クエリ、3回以上検索されクリックされたもの）／Abandoned Queries（クリック率が低い人気クエリ）／**No Results Queries（ゼロ件検索）**／Impression distribution
- 参照可能期間: 過去31日間（日次）または過去12か月（月次）。ダウンロード（Excel形式）でより広い期間を確認可能（31日分は日別タブ、12か月分は月別タブ）

**使用パーツ**: Microsoft Search の Insights（サイトコレクション単位の検索使用状況レポート）

**出典URL**:
- https://learn.microsoft.com/en-us/sharepoint/view-search-usage-reports-modern-sites（レポート種別・確認手順・期間の一次情報。全文確認済み）

**確度**: 高（一次情報に明記された機能であり、モダンサイトで標準提供されている）

**注意点**: 刷新後の「推移」で測るという考え方自体は妥当（ゼロ件検索率の改善などをbefore/afterで比較できる）が、**刷新前後でサイト構成・URL体系・インデックス範囲が変わる場合、クエリの母集団や意味が変化し単純比較が難しくなる可能性**がある点は付言すべき。また本レポートは**サイトコレクション管理者権限**が必要な点、保持期間が最大12か月である点（それ以前のデータはダウンロード保存していない限り参照不可）に注意。

---

## WF/回答への含意

1. **V8**: クイックリンクは4個以上の横並びが実務上可能だが、公式ドキュメントに固定列数の保証はない。「必ず◯列」という要件がある場合はプレビュー環境での実測を提案し、断定を避ける。
2. **V9**: 「サブブロックごとに別背景色」はセクション単位でしか公式サポートされない。要望があれば「セクション分割」を代替案として明示的に提示する。
3. **V10**: 5カテゴリの見せ方は「カテゴリごとに5つのHighlighted content web part＋各カテゴリ見出し」を標準案として推奨。1web partでのグループ表示機能は無い前提で設計する。
4. **V11**: クイックリンクの1タイル＝1リンクを前提に設計し、説明文＋複数リンクが必要な項目はテキストweb partで別途構成する。
5. **V12**: SharePoint agentの「外部リンクを探し当てて提示する」能力は一次情報で未確定。WFやクライアント提案では「対応する可能性が高いが未確証」と明記し、必要なら実機検証をタスク化する。
6. **V13**: 現行検索ログはベースラインとして使える（No Results Queries含む）。KPI設計は「刷新前後のゼロ件検索率・人気検索語の推移比較」を軸にでき、新規のログ基盤構築は不要。

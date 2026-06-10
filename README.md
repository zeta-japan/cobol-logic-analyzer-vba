# COBOL ロジック解析ツール

COBOLソースを解析し、Excel上で単体テスト設計を支援する。生成シート (ver3.0):

1. **COBOLソース** — 元コード + Excel側通し行番号 (各シートからハイパーリンク)
2. **ロジック階層(ソース順)** — IF / EVALUATE / SEARCH の入れ子をソース順のツリー表示
3. **ロジック階層(実行順展開)** — PERFORM を呼出箇所でインライン展開した実行順ツリー。
   「前へ／次へ」ボタンでテストケース毎に 決定パス・最終Action・実行行 を標記
4. **テストケース候補** — 実行流から生成したシナリオをステップフロー形式で列挙
   (正常系 = C1 分岐網羅の最少ケース、異常系 = ABEND 到達 + 外部 CALL 異常の合成)
5. **分岐カバレッジ表** — 行 = 検証Point (分岐アーム) × 列 = ケース のマトリクス。
   どのケースにも通らない行は赤 = 漏れ (恒真/恒偽分岐 = デッドコードもここに現れる)
6. **入出力-想定結果** — 正常系ケース毎の入力設定 (LINKAGE / DB前提 / サブ戻り値前提) と
   出力想定値 (最終代入)。実測値・判定の記入欄つき
7. **Driver雛形** — 正常系ケースを順次実行する Driver スケルトン (設定 → CALL → DISPLAY)
8. **分岐カバレッジ / 呼出関係 / 呼出関係図** — 分岐網羅集計・呼出関係表・呼出関係図

シングル `.xlsm` 内で完結し、外部ツール (.NET / Python 等) に依存しない。
社内および客先で **Excel + Windows 標準 COM のみ** が許可される環境向け。

---

## ファイル構成

```
cobol-logic-analyzer-vba/
├── README.md
├── LICENSE                              MIT
├── CobolAnalyzer.xlsm                   配布物 (Phase 5 以降に同梱)
├── vba/                                 VBA ソース (差分レビュー用)
│   ├── Main.bas                         エントリ
│   ├── CobolParser.bas                  解析エンジン本体
│   ├── CobolEncoding.bas                CP932 / UTF-8 自動判定読込
│   ├── JsonWriter.bas                   OrderedDict→JSON シリアライザ
│   ├── JsonParser.bas                   JSON パーサ + シートヘルパー
│   ├── CobolLogicViewer.bas             5 シート描画
│   ├── OrderedDict.cls                  挿入順保持の辞書クラス
│   ├── TestRunner.bas                   Assert ヘルパー + Run_All_Tests
│   └── Test_*.bas                       テスト群
└── samples/
    ├── input/                           ICASE1.cbl / ICASE2.cbl / sample_std.cbl / sample_prefixed.cbl
    └── golden/                          回帰テスト用 JSON fixture
```

### サンプル COBOL (samples/input)

| ファイル | 用途 |
|---|---|
| `ICASE1.cbl` | 基本サンプル。7段ネスト IF を含む保守的な構造 (回帰テストの基準) |
| `ICASE2.cbl` | **分岐構文 網羅サンプル**。下記の全構文 + 6段ネストをカバー。デモ/動作確認向け |
| `sample_std.cbl` | 標準書式 (プレフィックスなし) の判定確認用 |
| `sample_prefixed.cbl` | プレフィックス書式の自動判定確認用 |

`ICASE2.cbl` が網羅する構文 (機微な語は不使用・区分判定ドメイン):

- 制御構文: `IF / ELSE / END-IF`、`EVALUATE / WHEN / WHEN OTHER / END-EVALUATE`、
  `SEARCH / SEARCH ALL / AT END / WHEN / END-SEARCH`
- 呼出: `PERFORM`、`PERFORM ... THRU ...`、`CALL`、`GO TO`
- 命令: `MOVE / COMPUTE / READ / WRITE / REWRITE / DELETE`
- 警告誘発: `NEXT SENTENCE`、複合条件 `AND / OR`
- ネスト深さ **6 段** (`IF > IF > IF > EVALUATE > IF > IF`)
- 解析結果: branchCount=19 / pathCount=96 / 警告 6 種 / SEARCH ALL 検出

---

## アーキテクチャ

- **解析エンジン**: COBOL を読み、AST・実行パス・カバレッジ・呼出関係・元ソース を内部 JSON 構造で構築
- **可視化**: 同 JSON 構造を Excel に 5 シート描画
- **外部依存ゼロ**: `ADODB.Stream` / `Scripting.Dictionary` / `VBScript.RegExp` のみ (すべて Windows 標準同梱)

---

## セットアップ

1. 新規 Excel ブックを作成し `CobolAnalyzer.xlsm` として保存
2. VBE (Alt+F11) を開き、ツール → 参照設定 → **Microsoft Scripting Runtime** にチェック
3. ファイル → ファイルのインポート で `vba/` 配下の全ファイル (`.bas` / `.cls`) を取込
4. ブックを保存 (マクロ有効ブック `.xlsm` 形式)

---

## 使い方

### 初回セットアップ (配布前に 1 回だけ)

VBE のイミディエイトウィンドウで:

```
Main.SetupControlSheet
```

→ 「コントロール」シートと **「COBOL ソースを選択して解析」ボタン**が作られます。
このあとブックを保存すれば、ボタンは永続します。配布版はこの状態で渡します。

### 通常の使い方 (利用者)

1. 名前に **ABEND** を含む SECTION は自動で「異常終了扱い」になります (登録不要)。
   他の命名の ABEND 処理段 (例: `ERR-EXIT-SEC`) は「コントロール」シートの
   **B24〜B29** に登録してください。適用結果はテストケース候補シートの先頭に表示されます
2. 「コントロール」シートの **ボタンをクリック**
3. ファイル選択ダイアログから `.cbl` (または `.txt`) を選択
4. 解析結果が上記の各シートに生成される (再実行で自動上書き)

ボタンを使わず直接実行する場合: Excel で **Alt + F8 → `PickCobolAndBuild`**、または
イミディエイトで `Main.AnalyzeAndBuild "C:\path\to\sample.cbl"`。

中間 JSON は `%TEMP%` に一時生成され、描画後に**自動削除**されます。解析対象フォルダは変更しないため、利用者がファイルを掃除する必要はありません。

### 外部依存・セキュリティ

- ネットワーク通信なし / シェル実行 (WScript.Shell 等) なし / PowerShell・.NET 不使用
- 使用する COM は Windows 標準同梱のみ: `VBScript.RegExp` (正規表現)、`ADODB.Stream` (ファイル入出力)、`Scripting.Dictionary` (JSON パース補助)
- ファイル書き込みは `%TEMP%` の中間 JSON のみ (描画後に自動削除)。解析対象ソースは読み取り専用

### 回帰テスト

```
Run_All_Tests
```

→ ICASE1〜ICASE3 サンプルで各フェーズの主要数値 (branchCount / pathCount / マージ / 実行流ケース生成など) を検証。`TestResults` シートにも結果が出ます。

網羅サンプル `ICASE2.cbl` を実際に解析するには「コントロール」シートのボタン、または
イミディエイトで `Main.AnalyzeAndBuild ThisWorkbook.path & "\samples\input\ICASE2.cbl"`。

---

## ロードマップ

| Phase | 状態 | 目標 |
|---|---|---|
| 1. ハロー JSON | ✅ 完了 | プログラム名 + 行数だけの JSON 出力 |
| 2. AST + 構造抽出 | ✅ 完了 | `rootNodes` / `programStructure` |
| 3. 呼出グラフ + カバレッジ枠 | ✅ 完了 | `callGraph` / `coverage.branches` |
| 4. パス列挙 + テストケース | ✅ 完了 | `testCases` (パス上限 200 で打切) |
| 5. 5 シート描画 | ✅ 完了 | `CobolLogicViewer.BuildCobolReport` 結合 |
| 6. UI + 配布 | ✅ 完了 | ワンクリックボタン + コントロールシート (`SetupControlSheet`) |
| ver2.0 呼出関係図 / 入力項目 / Driver | ✅ 完了 | セルツリー呼出図・データ項目抽出・雛形生成 |
| ver2.1 実行順ツリー + ケース標記 | ✅ 完了 | 実行順展開 / 決定パス標記 / 実行行着色 |
| ver3.0 テストケース再設計 | ✅ 完了 | 実行流枚挙 + 定数伝播 + C1 最少ケース + 合成異常 + 分岐カバレッジ表 + 入出力-想定結果 + ケース別 Driver |

---

## ライセンス

[MIT](LICENSE)

# COBOL ロジック解析ツール

COBOLソースを解析し、Excel上で **5シート** にまとめて単体テスト設計を支援する。

1. **COBOLソース** — 元コード + Excel側通し行番号 (ロジック階層からハイパーリンク)
2. **ロジック階層** — IF / ELSE / EVALUATE / WHEN / SEARCH の入れ子構造をツリー表示
3. **テストケース候補** — 全実行パスを列挙して TC-001 形式のテストケース表を自動生成
4. **分岐カバレッジ** — 生成テストケースが各分岐を網羅しているかの集計
5. **呼出関係** — PERFORM / CALL / SECTION / 段落 の呼出関係表

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
    ├── input/                           ICASE1.cbl / sample_std.cbl / sample_prefixed.cbl
    └── golden/                          回帰テスト用 JSON fixture
```

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

## Phase 1: ハロー JSON (現状)

最小プロトタイプ。`samples/input/ICASE1.cbl` を読んでプログラム名と論理行数を JSON 化し A1 に出力する。

実行:

- VBE のイミディエイトウィンドウで `Sub_RunHello` を実行
- 期待出力: `{"summary":{"programName":"ICASE1","lines":NNN,"prefixDetected":true,"prefixStyle":"prefixed","prefixRatio":1}}`

テスト:

- イミディエイトで `Run_All_Tests` を実行
- 全 Assert が PASS、`TestResults` シートに結果が出る

---

## ロードマップ

| Phase | 状態 | 目標 |
|---|---|---|
| 1. ハロー JSON | 進行中 | プログラム名 + 行数だけの JSON 出力 |
| 2. AST + 構造抽出 | 未着手 | `rootNodes` / `programStructure` |
| 3. 呼出グラフ + カバレッジ枠 | 未着手 | `callGraph` / `coverage.branches` |
| 4. パス列挙 + テストケース | 未着手 | `testCases` (パス上限 200 で打切) |
| 5. 5 シート描画 | 未着手 | `CobolLogicViewer.BuildCobolReport` 結合 |
| 6. UI + 配布 | 未着手 | ワンクリックボタン |

---

## ライセンス

[MIT](LICENSE)

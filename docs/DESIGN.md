# 設計ドキュメント — Excel/CSV 月別売上集計レポート自動生成

> 将来の自分が「これは何だったか」を思い出すための設計メモ。
> README はユーザー向け、本書は開発者（=自分）向け。

最終更新: 2026-05-18

---

## 1. このプロジェクトの位置づけ

- ランサーズ営業用のポートフォリオ作品（受注用のサンプル兼デモ）
- 想定クライアント: 複数店舗の小売／飲食業で、毎月Excelで手集計している担当者
- 提供価値: POS/会計ソフトのCSVをフォルダに置く → exe実行 → KPI入りの月次Excelレポートが出来上がる、までを完全自動化

## 2. ゴールと非ゴール

### ゴール
- CSV → SQLite → Excel(KPI/サマリー/店舗詳細) の3ステップを1クリックで実行
- 非エンジニアが「exe1個 + CSV1個」で動かせる配布形態
- 年度・店舗を切り替えると該当シートだけ即座に差し替え（全体再生成しない）
- Excelネイティブのグラフ・条件付き書式を使い、開いた瞬間に経営判断ができる見た目

### 非ゴール
- 大規模データ（数百万行）対応 — pandas + SQLite で完結する規模を想定
- マルチユーザー同時編集、Web化、認証
- リアルタイム連携（CSV差分監視など）

## 3. アーキテクチャ全体像

```
[ data/sample/*.csv ]
        │  pandas read_csv（utf-8-sig, parse_dates）
        │  必須列バリデーション（日付/店舗名/売上金額/客数）
        ▼
[ data/sales.db (SQLite) ]   ←─ UPSERT:（日付×店舗名）一致行は上書き、それ以外は保持
        │
        │  load_from_db(year)  → pandas DataFrame
        │  build_monthly_pivot → 売上ピボット, 客数ピボット
        ▼
[ openpyxl で Workbook 構築 ]
   ├─ KPIダッシュボード（先頭シート、9枚のKPIカード + ミニテーブル + 凡例）
   ├─ 月別集計サマリー（全店一覧表 + 棒グラフ2種）
   └─ 店舗詳細（GUIで選択した店舗のみ、折れ線グラフ付き）
        ▼
[ output/monthly_report.xlsx ]
```

### 主要な分割

| レイヤ | ファイル | 役割 |
|---|---|---|
| GUI | `app/gui_launcher.pyw` | customtkinter のランチャー。年度/店舗の選択UIとボタン、起動時セットアップ |
| ドメイン | `app/generate_report.py` | CSV読込・DB入出力・ピボット集計・Excel生成すべて |
| 開発補助 | `scripts/generate_sample_data.py` | ポートフォリオ用ダミーCSV生成 |
| ビルド | `sales_report.spec` | PyInstaller の仕様（単一exe化） |

意図的にレイヤを薄く保っている（GUI と generate_report の2ファイルのみ）。
ポートフォリオ規模では over-engineering を避けたい。

## 4. データモデル

### CSV 入力フォーマット
```
日付,店舗名,売上金額,客数
2024-01-01,渋谷店,142000,75
```
- 文字コードは utf-8-sig を前提（Excel保存CSVのBOM対応）
- 日付は `YYYY-MM-DD`
- 列名は完全一致でバリデーション。違うとエラー停止（曖昧マッチはやらない）

### SQLite テーブル `sales`
| 列 | 型 | 備考 |
|---|---|---|
| 日付 | TEXT (YYYY-MM-DD) | CSV由来 |
| 店舗名 | TEXT | CSV由来 |
| 売上金額 | INTEGER | 円 |
| 客数 | INTEGER | 人 |
| 年 | INTEGER | 取り込み時に派生 |
| 月 | INTEGER | 取り込み時に派生 |

PRIMARY KEY は張っていない。UPSERTは pandas の merge + 差し替えで実装（`import_csv_to_db`）。
理由: ポートフォリオ規模では SQL の ON CONFLICT を使うより、pandasで `left_only` を取って `to_sql(if_exists="replace")` する方がコードが短く読める。

## 5. Excel生成の方針

### シート構成（順序固定）
1. **KPIダッシュボード** — 経営者向けのトップシート。KPIカード9枚＋月別ミニテーブル＋凡例。
2. **月別集計サマリー** — 全店舗×12ヶ月の一覧表＋全店合計棒グラフ＋店舗別比較棒グラフ。
3. **店舗詳細** — 1シートのみ（GUI選択中の店舗で差し替え）。月次表＋折れ線グラフ。

### 設計上のこだわり
- **数値の単位は動的に決める**: 100万円以上は「千円」、未満は「円」表示（`_yen_fmt`/`_yen_unit`）。実データ規模に応じて見やすい単位を選ぶ。
- **前月比は条件付き書式で色分け**: 赤(<-5%)/黄(±5%)/緑(>+5%)。Excel上で再計算しても色が追従する。
- **配色は Excel 標準カラー基準**: ヘッダー紺 `1F4E79`、サブ青 `2E75B6`、強調セル `BDD7EE`。WordやPowerPointと並べても違和感がない配色。
- **KPIカードは3行構成のセル結合で実装**: ラベル / 値（大字） / 補足説明。
- **グラフは openpyxl ネイティブ**: 画像ではなくExcelグラフなので、ユーザーがフィルタしても追従する。

### 既知の openpyxl の罠
- **`ExtendedProperties.Application` バグ**: 3.1.4+ で app.xml に `"Compatible / Openpyxl"` が入るとExcelがチャートを誤レンダリングする。
  → `generate_report.py` 冒頭で `ExtendedProperties.__init__` をモンキーパッチして `"Microsoft Excel"` に固定。
- **`load_workbook` で `plot_area.spPr` が失われる**: 店舗シート差し替え時に既存チャートの枠線が消える。
  → `_reapply_plot_area_borders` で再付与。

## 6. GUI の状態遷移

```
起動
  │
  ├─ DB+Excel両方ある  → セレクター読み込み → 「準備完了」
  └─ どちらか無い       → CSV取り込み → ベースExcel生成 → セレクター読み込み

年度変更    → ベースExcel再生成 + 選択中店舗の店舗詳細シート差し替え
店舗変更    → 店舗詳細シートのみ差し替え
DB更新ボタン → CSV再取り込み → セレクター再ロード
Excelを開く  → os.startfile で既定アプリで開く
```

- 全ての重い処理は `threading.Thread(daemon=True)` でバックグラウンド実行
- 処理中は操作ボタンを disabled に
- Excelを開いたまま再生成すると `PermissionError` → ユーザーにExcelを閉じるよう促す

## 7. exe配布（PyInstaller）

- `sales_report.spec` で onefile ビルド
- `_BASE` をフリーズ時は `sys.executable.parent`、開発時は `__file__.parent.parent` に切り替え
  → exe を任意フォルダに置いても、その隣の `data/` `output/` を見に行く
- ユーザーが用意するのは `sales_report.exe` + `data/sample/*.csv` の2つだけ

## 8. 「なぜそうしたか」メモ（将来の自分への伝言）

- **GUI に customtkinter を採用**: tkinter標準だと見た目が古く、ポートフォリオの第一印象が落ちる。PySide6 はexeが肥大化するためボツ。
- **DB を SQLite にした**: CSVを毎回再パースすると遅い & 履歴を保持できない。差分追加の挙動（CSVにない行は残す）は SQLite だと自然に書ける。
- **ピボットを pandas に任せた**: openpyxl で集計すると行ループになり可読性が落ちる。pandas で2行で済むものは pandas に寄せる。
- **店舗詳細シートを1枚だけにした**: 当初は全店舗を別シートに展開していた（`main()` には名残あり）が、GUIで切り替える前提なら1枚で十分。Excel容量も減る。
- **モンキーパッチを generate_report.py の上部に置いた**: 別モジュールに分けるほどでもなく、生成前に確実に読まれる位置で十分。コメントで理由を明示。

## 9. 拡張したくなったときの手がかり

| やりたいこと | どこを触る |
|---|---|
| 新しいKPIカード追加 | `write_kpi_sheet` の Row 7-13 ブロック |
| 新しい集計指標 | `build_monthly_pivot` を拡張 or 別ピボット関数を追加 |
| 列名のゆらぎ吸収 | `import_csv_to_db` の `REQUIRED` を rename マップに進化 |
| 別フォーマットへ出力（PDFなど） | `build_base_workbook` の出力部分を分岐 |
| 複数年比較シート | `write_*_sheet` 系を真似て新規追加。GUIにラジオ追加 |

## 10. 開発・ビルドコマンド早見

```powershell
# 開発実行
pip install -r requirements.txt
python app/gui_launcher.pyw

# サンプルCSV再生成
python scripts/generate_sample_data.py

# exeビルド
pyinstaller sales_report.spec
# → dist/sales_report.exe

# CLI単発実行（GUI無し）
python app/generate_report.py --input data/sample --output output/report.xlsx --year 2024
```

# xlsm2spec

**Excelマクロ資産（.xlsm / .xlam）から業務分析を行い、新システムの要求仕様書を生成する [opencode](https://opencode.ai) スキル**

ドキュメントも保守担当者も存在しない、古いExcelマクロ資産を解析して以下を明らかにします。

1. **業務分析** — ツールが何の業務を、どういう手順で、誰が、どう扱っているのか
2. **Excel特有の制約の読み解き** — Excel/マクロならではの制約の裏にある本質的な業務要件
3. **要求仕様** — 新システム（プログラム）として作り直す際の機能・データ・画面・帳票・運用要件

「コードの説明」ではなく「**業務を要件として再定義**」することを目的としたスキルです。

## 特徴

- **抽出**: VBA・シート構造・ボタン→マクロ割当・UserForm・DB参照をMarkdownへ自動抽出（`extract.py`）
- **業務分析ノウハウ**: 列マッピング（二次元配列→Range代入）からのDB↔シート復元、Null状態表現、業務ルール抽出、デッドコード・休眠機能の検出
- **制約パターンカタログ**: Excel資産に繰り返し現れる制約パターン（C1〜C23）を「現状→本質要件→新システムでの扱い」で整理
- **実DB検証**: Access DB（.accdb/.mdb）がある場合、実スキーマとの突き合わせで推測を確定
- **要求仕様化**: 機能・非機能・DB・画面・帳票・権限・移行の各要件と、ヒアリングが必要な「要確認事項」をレポート化

## 必要環境

- Python 3.8+
- 依存パッケージ:
  - `openpyxl` — Excelファイルの解析
  - `oletools` — VBAプロジェクト（vbaProject.bin）の解析
  - `access_parser` — Access DB（.accdb/.mdb）のスキーマ解析

```bash
pip install openpyxl oletools access_parser
```

## インストール（別PCでスキルとして利用する手順）

### 一発インストール（推奨）

```bash
curl -fsSL -H "Accept: application/vnd.github.raw" https://api.github.com/repos/kuwa2005/ExcelSpecification/contents/install.sh | bash
```

- スキルをグローバル（`~/.config/opencode/skills/`）に配置し、Python依存（openpyxl / oletools / access_parser）も導入します
- Access DB（.accdb/.mdb）連携ツールも**デフォルトで解析可能**です（`--db` で実スキーマを検証）

- 特定プロジェクトだけで使う場合は `--project <dir>` を付与

```bash
curl -fsSL -H "Accept: application/vnd.github.raw" https://api.github.com/repos/kuwa2005/ExcelSpecification/contents/install.sh | bash -s -- --project ./myproject
```

### 手動インストール

詳細は **[INSTALL.md](INSTALL.md)** を参照してください。

```bash
# 1. 取得
git clone https://github.com/kuwa2005/ExcelSpecification.git

# 2. グローバル配置（全プロジェクトで利用可）
mkdir -p ~/.config/opencode/skills
cp -r ExcelSpecification/.opencode/skills/xlsm2spec ~/.config/opencode/skills/

# 3. 依存導入
pip install openpyxl oletools access_parser

# 4. opencode を再起動 → 解析対象の .xlsm を渡すと自動発動
```

## 使い方

opencode のセッション内で、解析対象の `.xlsm` パスを伝えるとスキルが自動発動します。

```
例: 「./legacy_tool.xlsm の業務を分析して新システムの要求仕様書を作成して」
```

または、抽出のみを手動実行することもできます。

```bash
python3 .opencode/skills/xlsm2spec/scripts/extract.py <対象.xlsm> -o <作業ディレクトリ>
# Access DB連携ツールなら --db で実スキーマも検証
python3 .opencode/skills/xlsm2spec/scripts/extract.py <対象.xlsm> -o <作業ディレクトリ> --db data.accdb
```

### 生成される成果物

| 成果物 | 内容 |
|---|---|
| `00_workbook_overview.md` | ファイル情報・シート構成・図形数 |
| `10_sheet_list.md` | シート一覧と役割推定 |
| `20_vba_summary.md` | モジュール/プロシージャ一覧・呼び出し関係・DB参照・メッセージ |
| `25_forms.md` | UserFormのコントロール・種別・イベント・プロパティ |
| `30_buttons.md` | ボタン→マクロ→定義モジュール割当 |
| `40_cross_references.md` | VBA↔シート↔DBのクロス参照 |
| `50_db_schema.md` | Access DBのスキーマ検証とVBA参照の突合（`--db` 指定時のみ） |
| `sheets/*.md` | シートごとの列・数式・検証・コメント |
| `vba/*.txt` | VBAモジュールの完全ソース |

最終成果物として `<対象名>_業務分析_要求仕様.md`（業務分析 / 制約の読み解き / 要求仕様 / 要確認事項）を生成します。

## ディレクトリ構成

```
.
├── README.md
├── INSTALL.md                # 別PCでのインストール手順
└── .opencode/
    └── skills/
        └── xlsm2spec/          # このスキル本体
            ├── SKILL.md        # 解析ノウハウ・ワークフロー・制約パターンカタログ
            └── scripts/
                └── extract.py  # 抽出スクリプト
```

## ライセンス

このリポジトリに含まれるサンプル資産（example.xlsm / data.accdb）は外部に公開しません（.gitignore で除外）。スキル本体（SKILL.md / extract.py）は自由に利用してください。

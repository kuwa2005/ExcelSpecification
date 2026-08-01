# 別PCでのインストール手順（xlsm2spec）

このスキルを**別のPC**で使うための手順です。開いている端末でopencodeを起動したまま作業しないでください（設定は起動時に読み込まれます。配置後は**opencodeを再起動**します）。

## 0. 一発インストール（推奨）

下記の1行で、スキルの配置とPython依存の導入まで完了します。

```bash
curl -fsSL -H "Accept: application/vnd.github.raw" https://api.github.com/repos/kuwa2005/ExcelSpecification/contents/install.sh | bash
```

- スキルをグローバル（`~/.config/opencode/skills/`）に配置します
- openpyxl / oletools を自動導入します
- Access DB連携ツールを解析する場合は `--with-db` を付与:

```bash
curl -fsSL -H "Accept: application/vnd.github.raw" https://api.github.com/repos/kuwa2005/ExcelSpecification/contents/install.sh | bash -s -- --with-db
```

- 特定プロジェクトだけで使う場合:

```bash
curl -fsSL -H "Accept: application/vnd.github.raw" https://api.github.com/repos/kuwa2005/ExcelSpecification/contents/install.sh | bash -s -- --project ./myproject
```

インストールスクリプトのオプション:

| オプション | 説明 |
|---|---|
| `--project <dir>` | プロジェクト単位（`<dir>/.opencode/skills/`）でインストール |
| `--with-db` | Access DB解析用の `access_parser` も導入 |
| `--skip-deps` | Python依存の導入をスキップ |
| `--no-check` | import 検証をスキップ |

一発インストールを使わない場合の手動手順は以下を参照してください。

## 1. opencode のインストール（未導入の場合）

公式ドキュメント <https://opencode.ai> の手順に従って opencode をインストールします（既に導入済みの場合はスキップ）。

動作確認:

```bash
opencode --version
```

## 2. リポジトリの取得

```bash
git clone https://github.com/kuwa2005/ExcelSpecification.git
```

> このリポジトリは「opencodeスキル本体」のみを公開しています。解析対象のExcel資産（.xlsm）は含まれていません。

## 3. スキルの配置

opencode は `SKILL.md` を含むディレクトリをスキルとして自動認識します。配置場所は2通りあります。

| スコープ | パス | 有効範囲 |
|---|---|---|
| グローバル | `~/.config/opencode/skills/<名前>/SKILL.md` | 全プロジェクトで利用可能 |
| プロジェクト | `<プロジェクト>/.opencode/skills/<名前>/SKILL.md` | そのプロジェクトのみ |

### A. 全プロジェクトで使う場合（推奨）

```bash
mkdir -p ~/.config/opencode/skills
cp -r ExcelSpecification/.opencode/skills/xlsm2spec ~/.config/opencode/skills/
```

### B. 特定プロジェクトだけで使う場合

```bash
cd <解析したいプロジェクトのディレクトリ>
mkdir -p .opencode/skills
cp -r <クローン先>/ExcelSpecification/.opencode/skills/xlsm2spec .opencode/skills/
```

### C. リポジトリをそのままプロジェクトとして使う場合

リポジトリ内の `.opencode/skills/xlsm2spec/` は元から配置済みのため、コピー不要です。クローンしたディレクトリに移動してそのまま利用できます。

```bash
cd ExcelSpecification
```

## 4. Python依存パッケージの導入

```bash
pip install openpyxl oletools
```

- `oletools`（含む `olefile`）: VBAプロジェクト（vbaProject.bin）の解析に必須
- `openpyxl`: シート構造の解析に必須
- `access_parser`: **Access DB（.accdb/.mdb）連携ツール**を解析する場合のみ追加

```bash
pip install access_parser
```

検証:

```bash
python3 -c "import openpyxl, oletools"
```

## 5. 動作確認

配置後、**opencodeをいったん終了して再起動**してください（設定・スキルは起動時に読み込まれます）。

起動後にスキル一覧を確認します:

```
/help
```

`xlsm2spec`（または `xlsm2spec` の説明文）が表示されればインストール成功です。

## 6. 使い方

opencode のセッション内で、解析対象の `.xlsm` を渡すとスキルが自動発動します。

```
例: 「./legacy_tool.xlsm の業務を分析して、新システムの要求仕様書を作成して」
```

トリガー例: `xlsm` / `Excelマクロ` / `VBA解析` / `仕様書作成` / `仕様化` / `業務分析` / `要件定義` / `レガシーExcel資産` / `Access連携ツールの再構築`

抽出のみを手動で実行する場合:

```bash
python3 ~/.config/opencode/skills/xlsm2spec/scripts/extract.py <対象.xlsm> -o <出力ディレクトリ>
```

## 7. 更新方法

```bash
git -C ExcelSpecification pull
# グローバル配置の場合は再コピー
cp -r ExcelSpecification/.opencode/skills/xlsm2spec ~/.config/opencode/skills/
# opencode を再起動
```

## トラブルシューティング

| 症状 | 対処 |
|---|---|
| スキルが認識されない | `~/.config/opencode/skills/xlsm2spec/SKILL.md` の存在と、ファイル名が `SKILL.md`（大文字）であることを確認。opencode を完全終了して再起動 |
| `import openpyxl` が失敗する | `pip install openpyxl oletools` を再実行。環境が externally-managed の場合は `--break-system-packages` を付与 |
| 抽出スクリプトがエラー | `python3 -c "import openpyxl, oletools"` で依存を確認。VBA解析は `oletools` が必須 |
| DB（.accdb）が読めない | `pip install access_parser` を実行 |

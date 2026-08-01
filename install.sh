#!/usr/bin/env bash
#
# xlsm2spec skill installer
#
#   curl -fsSL -H "Accept: application/vnd.github.raw" https://api.github.com/repos/kuwa2005/ExcelSpecification/contents/install.sh | bash
#
# options:
#   --project <dir>   プロジェクト単位でインストール (<dir>/.opencode/skills)
#   --skip-deps       Python 依存の導入をスキップ
#   --no-check        検証 (import 確認) をスキップ
#
set -euo pipefail

REPO_URL="https://github.com/kuwa2005/ExcelSpecification.git"
SKILL_REL=".opencode/skills/xlsm2spec"
CORE_DEPS="openpyxl oletools access_parser"
SKIP_DEPS=0
NO_CHECK=0
TARGET="global"

usage() {
  sed -n '2,12p' "$0"
}

# ---- 引数解析 ----
while [ $# -gt 0 ]; do
  case "$1" in
    --project)
      [ $# -ge 2 ] || { echo "--project にはディレクトリが必要です"; exit 1; }
      TARGET="$2"
      shift 2
      ;;
    --skip-deps)  SKIP_DEPS=1; shift ;;
    --no-check)   NO_CHECK=1; shift ;;
    -h|--help)    usage; exit 0 ;;
    *)            echo "不明な引数: $1"; usage; exit 1 ;;
  esac
done

echo "== xlsm2spec skill installer =="

# ---- 必須コマンド ----
command -v git  >/dev/null 2>&1 || { echo "[NG] git がインストールされていません"; exit 1; }
command -v python3 >/dev/null 2>&1 || { echo "[NG] python3 がインストールされていません"; exit 1; }

# ---- 配置先 ----
if [ "$TARGET" = "global" ]; then
  CONFIG_HOME="${XDG_CONFIG_HOME:-$HOME/.config}"
  DEST="$CONFIG_HOME/opencode/skills/xlsm2spec"
else
  DEST="$TARGET/.opencode/skills/xlsm2spec"
fi

# ---- スキル取得（一時クローン → コピー）----
echo "[1/3] スキルを取得中..."
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
if ! git clone --depth 1 --quiet "$REPO_URL" "$TMP/repo" 2>/dev/null; then
  echo "[NG] リポジトリの取得に失敗しました"
  exit 1
fi
if [ ! -d "$TMP/repo/$SKILL_REL" ]; then
  echo "[NG] リポジトリ内にスキルが見つかりません"
  exit 1
fi
mkdir -p "$(dirname "$DEST")"
rm -rf "$DEST"
cp -r "$TMP/repo/$SKILL_REL" "$DEST"
echo "     配置先: $DEST"

# ---- Python 依存 ----
PY3="$(command -v python3)"
VENV_PY=""

# pip を介した導入を試行する（PEP668環境は --break-system-packages にフォールバック）
pip_install() {
  if "$PY3" -m pip install --quiet "$@" 2>/dev/null; then
    return 0
  fi
  "$PY3" -m pip install --quiet --break-system-packages "$@" 2>/dev/null
}

DEPS="$CORE_DEPS"
if [ "$SKIP_DEPS" = "1" ]; then
  echo "[2/3] 依存の導入をスキップ"
else
  echo "[2/3] Python 依存を導入中: $DEPS"
  DEP_OK=0

  # 1) 既存の pip
  if "$PY3" -m pip --version >/dev/null 2>&1; then
    if pip_install $DEPS; then DEP_OK=1; fi
  fi

  # 2) pip 自体が無い → get-pip.py でユーザー領域にブートストラップ
  if [ "$DEP_OK" = "0" ] && command -v curl >/dev/null 2>&1; then
    echo "     pip が見つかりません → get-pip.py でブートストラップを試行"
    GETTMP="$(mktemp -d)"
    if curl -fsSL https://bootstrap.pypa.io/get-pip.py -o "$GETTMP/get-pip.py" 2>/dev/null \
       && "$PY3" "$GETTMP/get-pip.py" --user --break-system-packages >/dev/null 2>&1 \
       && pip_install --user $DEPS; then
      DEP_OK=1
    fi
    rm -rf "$GETTMP"
  fi

  # 3) ユーザー領域も不可 → スキル内 .venv を作成して導入（extract.py が自動利用）
  if [ "$DEP_OK" = "0" ]; then
    echo "     ユーザー領域への導入に失敗 → スキル内 .venv を試行"
    if "$PY3" -m venv "$DEST/.venv" >/dev/null 2>&1; then
      if "$DEST/.venv/bin/python3" -m pip install --quiet $DEPS 2>/dev/null; then
        DEP_OK=1
        VENV_PY="$DEST/.venv/bin/python3"
        echo "     導入先: $DEST/.venv（extract.py が自動で使用します）"
      else
        rm -rf "$DEST/.venv"
      fi
    fi
  fi

  if [ "$DEP_OK" != "1" ]; then
    echo "[NG] 依存の導入に失敗しました"
    echo "     sudo apt install python3-pip python3-venv を導入して再実行してください"
    echo "       （または手動で: python3 -m pip install openpyxl oletools access_parser）"
    exit 1
  fi
fi

# ---- 検証 ----
VCHECK="${VENV_PY:-$PY3}"
if [ "$NO_CHECK" = "1" ]; then
  echo "[3/3] 検証をスキップ"
else
  echo "[3/3] 検証中..."
  if "$VCHECK" -c "import openpyxl, oletools" 2>/dev/null; then
    echo "     openpyxl / oletools: OK"
  else
    echo "     [warn] import に失敗しました（依存導入の失敗が考えられます）"
    "$VCHECK" -m pip list 2>/dev/null | grep -iE "openpyxl|oletools" || true
  fi
  if "$VCHECK" -c "import access_parser" 2>/dev/null; then
    echo "     access_parser: OK"
  else
    echo "     [warn] access_parser の import に失敗しました"
  fi
fi

# ---- 完了 ----
echo ""
echo "== インストール完了 =="
if command -v opencode >/dev/null 2>&1; then
  echo "  次回以降の opencode で 'xlsm2spec' スキルが利用できます。"
  echo "  ※現在起動中の opencode は終了して再起動してください。"
else
  echo "  opencode が未インストールです。 https://opencode.ai から導入してください。"
fi
echo "  ※同じコマンドを再実行すると最新版にアップデートされます"

cat <<'TUTORIAL'

■ はじめに（簡単な使い方）
────────────────────────────────────────
① 解析したい .xlsm があるディレクトリで opencode を起動
      cd <対象のフォルダ>
      opencode

② 起動したら、解析したいファイルを伝えるだけでスキルが自動発動します
      「example.xlsm の業務を分析して、新システムの要求仕様書を作成して」

③ 出力されるもの
  ・ 抽出成果物（解析材料）: カレントディレクトリの ./.tmp/
      00_workbook_overview.md, 10_sheet_list.md, 20_vba_summary.md,
      25_forms.md, 30_buttons.md, 40_cross_references.md,
      sheets/*.md, vba/*.txt
  ・ 最終レポート: 対象ファイルと同じ場所に
      <対象名>_業務分析_要求仕様.md
      （業務フロー / データモデル / Excelの制約の読み解き /
        機能・非機能・DB・画面・帳票要件 / 要確認事項）

④ 抽出だけを手動で行う場合
      mkdir -p ./.tmp
      python3 <スキル>/scripts/extract.py <対象.xlsm> -o ./.tmp

⑤ Access DB (.accdb) 連携ツールもデフォルトで解析可能（access_parser 同梱）
      抽出時に --db で実スキーマも検証:
      python3 <スキル>/scripts/extract.py <対象.xlsm> -o ./.tmp --db data.accdb
────────────────────────────────────────
TUTORIAL

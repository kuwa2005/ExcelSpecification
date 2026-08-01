#!/usr/bin/env bash
#
# xlsm2spec skill installer
#
#   curl -fsSL https://raw.githubusercontent.com/kuwa2005/ExcelSpecification/main/install.sh | bash
#
# options:
#   --project <dir>   プロジェクト単位でインストール (<dir>/.opencode/skills)
#   --with-db         Access DB 解析用の access_parser も導入
#   --skip-deps       Python 依存の導入をスキップ
#   --no-check        検証 (import 確認) をスキップ
#
set -euo pipefail

REPO_URL="https://github.com/kuwa2005/ExcelSpecification.git"
SKILL_REL=".opencode/skills/xlsm2spec"
CORE_DEPS="openpyxl oletools"
DB_DEPS="access_parser"
WITH_DB=0
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
    --with-db)    WITH_DB=1; shift ;;
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
DEPS="$CORE_DEPS"
if [ "$WITH_DB" = "1" ]; then
  DEPS="$DEPS $DB_DEPS"
fi
if [ "$SKIP_DEPS" = "1" ]; then
  echo "[2/3] 依存の導入をスキップ"
else
  echo "[2/3] Python 依存を導入中: $DEPS"
  if ! python3 -m pip install --quiet $DEPS 2>/dev/null; then
    echo "     標準インストールに失敗 → --break-system-packages で再試行"
    python3 -m pip install --quiet --break-system-packages $DEPS || {
      echo "[NG] 依存の導入に失敗しました"
      echo "     sudo apt install python3-pip などで pip を準備して再実行してください"
      exit 1
    }
  fi
fi

# ---- 検証 ----
if [ "$NO_CHECK" = "1" ]; then
  echo "[3/3] 検証をスキップ"
else
  echo "[3/3] 検証中..."
  if python3 -c "import openpyxl, oletools" 2>/dev/null; then
    echo "     openpyxl / oletools: OK"
  else
    echo "     [warn] import に失敗しました（依存導入の失敗が考えられます）"
    python3 -m pip list 2>/dev/null | grep -iE "openpyxl|oletools" || true
  fi
  if [ "$WITH_DB" = "1" ]; then
    if python3 -c "import access_parser" 2>/dev/null; then
      echo "     access_parser: OK"
    else
      echo "     [warn] access_parser の import に失敗しました"
    fi
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

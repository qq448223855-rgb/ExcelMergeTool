#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PYTHON="${PYTHON:-python3}"
APP_NAME="Excel合并工具V1.0.1.app"
ZIP_NAME="Excel合并工具V1.0.1-macOS-arm64.zip"
DIST_DIR="$PROJECT_ROOT/dist"
BUILD_DIR="$PROJECT_ROOT/build"
RELEASE_DIR="$PROJECT_ROOT/release"

rm -rf "$DIST_DIR" "$BUILD_DIR" "$RELEASE_DIR"
mkdir -p "$RELEASE_DIR"

"$PYTHON" -m PyInstaller \
  --noconfirm \
  --clean \
  --distpath "$DIST_DIR" \
  --workpath "$BUILD_DIR" \
  "$PROJECT_ROOT/packaging/ExcelMergeTool.spec"

APP_PATH="$DIST_DIR/$APP_NAME"
if [[ ! -d "$APP_PATH" ]]; then
  echo "Build failed: $APP_PATH was not created." >&2
  exit 1
fi

/usr/bin/codesign --force --deep --sign - "$APP_PATH"
/usr/bin/codesign --verify --deep --strict "$APP_PATH"

cp -R "$APP_PATH" "$RELEASE_DIR/$APP_NAME"
(
  cd "$RELEASE_DIR"
  COPYFILE_DISABLE=1 /usr/bin/zip -9 -y -q -r "$ZIP_NAME" "$APP_NAME" \
    -x "*.DS_Store" "*/__MACOSX/*"
)

APP_BYTES="$(/usr/bin/du -sk "$RELEASE_DIR/$APP_NAME" | awk '{print $1 * 1024}')"
ZIP_BYTES="$(/usr/bin/stat -f '%z' "$RELEASE_DIR/$ZIP_NAME")"
APP_LIMIT=$((60 * 1024 * 1024))
ZIP_LIMIT=$((20 * 1024 * 1024))

echo "App size: $((APP_BYTES / 1024 / 1024)) MB"
echo "Zip size: $((ZIP_BYTES / 1024 / 1024)) MB"

if (( APP_BYTES > APP_LIMIT )); then
  echo "App exceeds the 60 MB release limit." >&2
  exit 1
fi

if (( ZIP_BYTES > ZIP_LIMIT )); then
  echo "Zip exceeds the 20 MB release limit." >&2
  exit 1
fi

echo "Release artifacts are in $RELEASE_DIR"

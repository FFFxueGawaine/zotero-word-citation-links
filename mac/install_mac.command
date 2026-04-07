#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SOURCE_TEMPLATE="$SCRIPT_DIR/Zotero.dotm"
TARGET_DIR="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word"
TARGET_TEMPLATE="$TARGET_DIR/Zotero.dotm"
BACKUP_DIR="$SCRIPT_DIR/backup"
LATEST_BACKUP_FILE="$BACKUP_DIR/LATEST_BACKUP.txt"
TIMESTAMP="$(date +"%Y%m%d-%H%M%S")"

print_line() {
  printf '%s\n' "$1"
}

pause_and_exit() {
  print_line ""
  read -r -p "Press Enter to close..."
}

is_word_running() {
  pgrep -x "Microsoft Word" >/dev/null 2>&1 || pgrep -f "Microsoft Word.app" >/dev/null 2>&1
}

print_line "Zotero Word Citation Links - Mac Installer"
print_line ""

if [[ ! -f "$SOURCE_TEMPLATE" ]]; then
  print_line "Error: Zotero.dotm was not found next to install_mac.command."
  pause_and_exit
  exit 1
fi

if is_word_running; then
  print_line "Please quit Microsoft Word before installing."
  pause_and_exit
  exit 1
fi

mkdir -p "$TARGET_DIR"
mkdir -p "$BACKUP_DIR"

if [[ -f "$TARGET_TEMPLATE" ]]; then
  BACKUP_TEMPLATE="$BACKUP_DIR/Zotero.dotm.backup.$TIMESTAMP.dotm"
  cp "$TARGET_TEMPLATE" "$BACKUP_TEMPLATE"
  printf '%s\n' "$BACKUP_TEMPLATE" > "$LATEST_BACKUP_FILE"
  print_line "Backup created:"
  print_line "$BACKUP_TEMPLATE"
  print_line ""
fi

cp "$SOURCE_TEMPLATE" "$TARGET_TEMPLATE"

print_line "Install finished."
print_line "Installed template:"
print_line "$TARGET_TEMPLATE"
print_line ""
print_line "Reopen Word and check the Zotero tab for:"
print_line "- Create Citation Links"
print_line "- Remove Citation Links"
print_line "- Zotero Citation Link (document character style)"

pause_and_exit

#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TARGET_DIR="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word"
TARGET_TEMPLATE="$TARGET_DIR/Zotero.dotm"
BACKUP_DIR="$SCRIPT_DIR/backup"
LATEST_BACKUP_FILE="$BACKUP_DIR/LATEST_BACKUP.txt"

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

resolve_backup() {
  if [[ -f "$LATEST_BACKUP_FILE" ]]; then
    local latest_backup
    latest_backup="$(cat "$LATEST_BACKUP_FILE")"
    if [[ -f "$latest_backup" ]]; then
      printf '%s\n' "$latest_backup"
      return 0
    fi
  fi

  local newest_backup
  newest_backup="$(ls -1t "$BACKUP_DIR"/Zotero.dotm.backup.*.dotm 2>/dev/null | head -n 1 || true)"
  if [[ -n "$newest_backup" && -f "$newest_backup" ]]; then
    printf '%s\n' "$newest_backup"
    return 0
  fi

  return 1
}

print_line "Zotero Word Citation Links - Mac Restore"
print_line ""

if is_word_running; then
  print_line "Please quit Microsoft Word before restoring."
  pause_and_exit
  exit 1
fi

if [[ ! -d "$BACKUP_DIR" ]]; then
  print_line "No backup directory was found."
  pause_and_exit
  exit 1
fi

BACKUP_TEMPLATE="$(resolve_backup || true)"
if [[ -z "$BACKUP_TEMPLATE" ]]; then
  print_line "No backup Zotero.dotm was found."
  pause_and_exit
  exit 1
fi

mkdir -p "$TARGET_DIR"
cp "$BACKUP_TEMPLATE" "$TARGET_TEMPLATE"

print_line "Restore finished."
print_line "Restored from:"
print_line "$BACKUP_TEMPLATE"
print_line ""
print_line "Reopen Word to verify the original Zotero template is active."

pause_and_exit

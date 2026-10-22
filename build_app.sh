#!/bin/zsh
# Builds ~/Desktop/PPTX to Markdown.app: an Automator droplet (cloned from the existing
# "Convert to Markdown.app") whose Run Shell Script action calls pptx2md.py on dropped files.
set -euo pipefail

HERE="${0:A:h}"
TEMPLATE="$HOME/Desktop/Convert to Markdown.app"
APP="$HOME/Desktop/PPTX to Markdown.app"
PY="$HERE/.venv/bin/python"

[[ -d "$TEMPLATE" ]] || { echo "missing template app: $TEMPLATE" >&2; exit 1; }
[[ -x "$PY" ]] || { echo "missing venv: run: /opt/homebrew/bin/python3 -m venv .venv && .venv/bin/pip install python-pptx pillow lxml pypdfium2" >&2; exit 1; }

rm -rf "$APP"
ditto "$TEMPLATE" "$APP"
xattr -cr "$APP"

COMMAND="\"$PY\" \"$HERE/pptx2md.py\" \"\$@\"
exit 0"
plutil -replace actions.0.action.ActionParameters.COMMAND_STRING -string "$COMMAND" "$APP/Contents/document.wflow"

PLIST="$APP/Contents/Info.plist"
plutil -replace CFBundleName -string "PPTX to Markdown" "$PLIST"
plutil -replace CFBundleIdentifier -string "com.apple.automator.PPTX-to-Markdown" "$PLIST"
plutil -replace CFBundleIconFile -string "AppIcon" "$PLIST"
plutil -remove CFBundleIconName "$PLIST" 2>/dev/null || true
plutil -replace NSAppleEventsUsageDescription -string \
  "PPTX to Markdown uses Microsoft PowerPoint to render full-slide snapshots and to open older .ppt files." "$PLIST"

"$PY" "$HERE/make_icon.py" "$APP/Contents/Resources/AppIcon.icns"

codesign --force --deep --sign - "$APP"
touch "$APP"
echo "built $APP"

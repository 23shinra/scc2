#!/bin/bash
# Install / uninstall daily launchd job (macOS, 08:00 local time)
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
LABEL="com.scc2.rfc-daily"
PLIST="$HOME/Library/LaunchAgents/${LABEL}.plist"
TEMPLATE="$ROOT/com.scc2.rfc-daily.plist.template"

chmod +x "$ROOT/run_daily.sh" "$ROOT/pdf_to_excel.py" "$ROOT/sync_rfc.py"

case "${1:-install}" in
  install)
    sed "s|__ROOT__|${ROOT}|g" "$TEMPLATE" >"$PLIST"
    launchctl bootout "gui/$(id -u)/${LABEL}" 2>/dev/null || true
    launchctl bootstrap "gui/$(id -u)" "$PLIST"
    launchctl enable "gui/$(id -u)/${LABEL}"
    echo "Установлено: каждый день в 08:00"
    echo "Плист: $PLIST"
    echo "Проверка сейчас: $ROOT/run_daily.sh"
    ;;
  uninstall)
    launchctl bootout "gui/$(id -u)/${LABEL}" 2>/dev/null || true
    rm -f "$PLIST"
    echo "Снято с автозапуска"
    ;;
  *)
    echo "Usage: $0 [install|uninstall]"
    exit 1
    ;;
esac

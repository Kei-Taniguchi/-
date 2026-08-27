#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."
PID_FILE=/tmp/household-budget-web-app.pid
LOG_FILE=/tmp/household-budget-web-app.log

# Stop only the server started by this app, if one is already running.
if [ -f "$PID_FILE" ]; then
  OLD_PID=$(cat "$PID_FILE" 2>/dev/null || true)
  if [ -n "${OLD_PID:-}" ] && kill -0 "$OLD_PID" 2>/dev/null; then
    kill "$OLD_PID" 2>/dev/null || true
    sleep 1
  fi
fi

# Dependencies are normally installed by postCreateCommand. Install if a
# Codespace was restored without node_modules.
if [ ! -x node_modules/.bin/http-server ]; then
  npm install
fi

nohup npm start > "$LOG_FILE" 2>&1 &
echo $! > "$PID_FILE"

# Wait until the HTTP endpoint actually responds. This avoids opening the
# forwarded browser port before the server is ready.
for i in {1..30}; do
  if node -e "require('http').get('http://127.0.0.1:8080/', r => process.exit(r.statusCode < 500 ? 0 : 1)).on('error', () => process.exit(1))"; then
    exit 0
  fi
  sleep 1
done

echo "Web server did not become ready on port 8080. Log: $LOG_FILE" >&2
cat "$LOG_FILE" >&2 || true
exit 1

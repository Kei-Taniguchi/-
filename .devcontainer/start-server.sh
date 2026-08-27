#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."

# Avoid duplicate preview servers when Codespaces restarts the container.
pkill -f "http-server .* -p 8080" 2>/dev/null || true

# Dependencies are normally installed by postCreateCommand. Install if a
# Codespace was restored without node_modules.
if [ ! -x node_modules/.bin/http-server ]; then
  npm install
fi

nohup npm start > /tmp/household-budget-web-app.log 2>&1 &
echo $! > /tmp/household-budget-web-app.pid

# Wait briefly for the listener so the forwarded port is ready before the
# Codespaces browser preview opens.
for i in {1..30}; do
  if (curl -fsS http://127.0.0.1:8080/ >/dev/null 2>&1); then
    exit 0
  fi
  sleep 1
done

echo "Web server did not become ready on port 8080. Log: /tmp/household-budget-web-app.log" >&2
cat /tmp/household-budget-web-app.log >&2 || true
exit 1

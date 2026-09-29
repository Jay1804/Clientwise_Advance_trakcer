#!/bin/bash
# One-shot deploy of a Streamlit app on the server (no Docker, no sudo).
# Usage:  bash deploy_app.sh <git_repo_url> <port>
# Example: bash deploy_app.sh https://github.com/Jay1804/Clientwise_Advance_trakcer 8502
set -e

REPO_URL="$1"
PORT="$2"
if [ -z "$REPO_URL" ] || [ -z "$PORT" ]; then
  echo "Usage: bash deploy_app.sh <git_repo_url> <port>"; exit 1
fi

REPO_NAME=$(basename "$REPO_URL" .git)
BASE="$HOME/apps"
D="$BASE/$REPO_NAME"

echo "== 1. Clone / update repo"
mkdir -p "$BASE"
if [ -d "$D/.git" ]; then git -C "$D" pull; else git clone "$REPO_URL" "$D"; fi
cd "$D"

echo "== 2. Python 3.11 (installed once via uv)"
if [ ! -x "$HOME/.local/bin/uv" ]; then
  /usr/local/bin/python3.8 -m pip install --user -U pip
  /usr/local/bin/python3.8 -m pip install --user uv
fi
"$HOME/.local/bin/uv" python install 3.11
PY=$("$HOME/.local/bin/uv" python find 3.11)

echo "== 3. Virtual environment + requirements"
[ -d venv ] || "$PY" -m venv venv
venv/bin/pip install --upgrade pip
# Old OS (glibc 2.17): pandas 2.3.3 has no wheel, use 2.3.2
sed 's/^pandas==2.3.3$/pandas==2.3.2/' requirements.txt > /tmp/req_server.txt
venv/bin/pip install -r /tmp/req_server.txt

echo "== 4. .env check"
if [ ! -f .env ]; then
  echo "!! No .env found in $D - create it (nano .env) with DB_HOST, DB_NAME, DB_USER, DB_PASSWORD, SMTP_* and re-run."
  exit 1
fi

echo "== 5. Watchdog script"
cat > start_app.sh <<EOS
#!/bin/bash
D=$D
cd "\$D" || exit 1
if ! pgrep -f "streamlit run app.py --server.port $PORT" > /dev/null; then
  nohup "\$D/venv/bin/streamlit" run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true >> "\$D/app.log" 2>&1 < /dev/null &
fi
EOS
chmod +x start_app.sh

echo "== 6. Cron (start at boot + check every minute)"
( crontab -l 2>/dev/null | grep -v "$D/start_app.sh"
  echo "@reboot $D/start_app.sh"
  echo "* * * * * $D/start_app.sh" ) | crontab -

echo "== 7. Start and verify"
./start_app.sh
sleep 6
pgrep -af "streamlit run app.py --server.port $PORT" | cut -c1-90
echo -n "Health: "; curl -s -m 5 "http://localhost:$PORT/_stcore/health"; echo
echo "Open: http://$(hostname -I | awk '{print $1}'):$PORT"

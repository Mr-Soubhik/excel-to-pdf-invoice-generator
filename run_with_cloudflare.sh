#!/bin/bash
# run_with_cloudflare.sh - Launch Web App + Free Cloudflare Tunnel

echo "Starting Invoice Web Application..."
./.venv/bin/python app.py &
APP_PID=$!

echo "Waiting for app to initialize..."
sleep 2

if ! command -v cloudflared &> /dev/null; then
    echo "Downloading cloudflared..."
    wget -q https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64 -O ./cloudflared
    chmod +x ./cloudflared
    CLOUDFLARED_CMD="./cloudflared"
else
    CLOUDFLARED_CMD="cloudflared"
fi

echo "======================================================="
echo "  CREATING FREE CLOUDFLARE PUBLIC TUNNEL FOR CLIENT  "
echo "======================================================="
$CLOUDFLARED_CMD tunnel --url http://127.0.0.1:5000

kill $APP_PID 2>/dev/null

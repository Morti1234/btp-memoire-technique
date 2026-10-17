#!/bin/bash

# Kill any existing streamlit or cloudflared
pkill -f "streamlit run"
pkill -f "cloudflared tunnel"

# Start streamlit in the background
nohup streamlit run app.py --server.port=8501 --server.address=0.0.0.0 > streamlit.log 2>&1 &
echo "Streamlit started."

# Wait a few seconds for it to boot
sleep 3

# Start cloudflared
rm -f cloudflared.log
nohup cloudflared tunnel --url http://localhost:8501 > cloudflared.log 2>&1 &
echo "Cloudflared started."

# Wait for the trycloudflare URL to appear in the logs
echo "Waiting for Cloudflare Tunnel URL..."
URL=""
for i in {1..30}; do
    URL=$(grep -o 'https://[a-zA-Z0-9-]*\.trycloudflare\.com' cloudflared.log | head -1)
    if [ ! -z "$URL" ]; then
        echo "FOUND_URL: $URL"
        break
    fi
    sleep 1
done

if [ -z "$URL" ]; then
    echo "Could not find URL. Logs:"
    cat cloudflared.log
fi

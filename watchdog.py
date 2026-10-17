import subprocess
import time
import os
import re
import urllib.request
import urllib.error

DESKTOP_PATH = os.path.expanduser("~/Desktop")
HTML_SHORTCUT = os.path.join(DESKTOP_PATH, "Accéder à MarketBuilding.html")
APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(APP_DIR, "watchdog.log")

def log(msg):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}\n"
    print(line, end="")
    with open(LOG_FILE, "a") as f:
        f.write(line)

def is_process_running(name):
    try:
        out = subprocess.check_output(["pgrep", "-f", name]).decode()
        return len(out.strip()) > 0
    except subprocess.CalledProcessError:
        return False

def check_url(url):
    try:
        req = urllib.request.Request(
            url, 
            headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=5) as response:
            return response.status in [200, 302, 303, 307]
    except Exception:
        return False

def get_current_tunnel_url():
    cloudflared_log = os.path.join(APP_DIR, "cloudflared.log")
    if not os.path.exists(cloudflared_log):
        return None
    try:
        with open(cloudflared_log, "r") as f:
            lines = f.readlines()
        for line in reversed(lines):
            match = re.search(r'https://[a-zA-Z0-9-]+\.trycloudflare\.com', line)
            if match:
                return match.group(0)
    except Exception:
        pass
    return None

def update_desktop_shortcut(tunnel_url):
    html_content = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="refresh" content="2;url={tunnel_url}">
    <title>Connexion à MarketBuilding...</title>
    <style>
        body {{
            background-color: #0e1117;
            color: #ffffff;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            height: 100vh;
            margin: 0;
            text-align: center;
        }}
        .card {{
            background: #1e293b;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.5);
            max-width: 500px;
        }}
        h2 {{ color: #3B82F6; margin-top: 0; }}
        a.btn {{
            display: inline-block;
            background: #2563EB;
            color: #ffffff;
            text-decoration: none;
            padding: 12px 24px;
            border-radius: 8px;
            font-weight: bold;
            margin: 10px 5px;
        }}
        a.btn:hover {{ background: #1D4ED8; }}
        a.btn-sec {{
            background: #334155;
        }}
        a.btn-sec:hover {{ background: #475569; }}
        .loader {{
            border: 4px solid #334155;
            border-top: 4px solid #3B82F6;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }}
        @keyframes spin {{ 0% {{ transform: rotate(0deg); }} 100% {{ transform: rotate(360deg); }} }}
    </style>
</head>
<body>
    <div class="card">
        <h2>🏗️ MarketBuilding Pro</h2>
        <div class="loader"></div>
        <p>Redirection automatique vers votre site en cours...</p>
        <div style="margin-top: 20px;">
            <a class="btn" href="{tunnel_url}">Accès en Ligne (Cloudflare)</a>
            <a class="btn btn-sec" href="http://localhost:8501">Accès Direct Local (Mac)</a>
        </div>
        <p style="font-size: 11px; color: #94a3b8; margin-top: 15px;">Mot de passe : <b>admin</b></p>
    </div>
</body>
</html>
"""
    try:
        with open(HTML_SHORTCUT, "w", encoding="utf-8") as f:
            f.write(html_content)
        log(f"Raccourci Bureau mis a jour: {tunnel_url}")
    except Exception as e:
        log(f"Erreur ecriture raccourci bureau: {e}")

def start_streamlit():
    log("Demarrage de Streamlit...")
    subprocess.Popen(
        ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"],
        cwd=APP_DIR,
        stdout=open(os.path.join(APP_DIR, "streamlit.log"), "a"),
        stderr=subprocess.STDOUT
    )
    time.sleep(3)

def restart_cloudflared():
    log("Redemarrage du tunnel Cloudflare...")
    subprocess.call(["pkill", "-f", "cloudflared tunnel"])
    time.sleep(1)
    subprocess.Popen(
        ["cloudflared", "tunnel", "--url", "http://localhost:8501"],
        cwd=APP_DIR,
        stdout=open(os.path.join(APP_DIR, "cloudflared.log"), "w"),
        stderr=subprocess.STDOUT
    )
    # Wait for new URL
    for _ in range(15):
        time.sleep(1)
        url = get_current_tunnel_url()
        if url:
            log(f"Nouveau tunnel detecte: {url}")
            update_desktop_shortcut(url)
            return url
    return None

def main():
    log("Watchdog MarketBuilding demarre.")
    while True:
        # 1. Verifier Streamlit
        if not check_url("http://localhost:8501"):
            log("Streamlit injoignable sur localhost:8501. Relance...")
            subprocess.call(["pkill", "-f", "streamlit run"])
            start_streamlit()

        # 2. Verifier Cloudflared
        url = get_current_tunnel_url()
        tunnel_ok = False
        if url and is_process_running("cloudflared tunnel"):
            tunnel_ok = check_url(url)

        if not tunnel_ok:
            log(f"Tunnel Cloudflare HS ou expire (URL: {url}). Relance...")
            new_url = restart_cloudflared()
            if new_url:
                log(f"Tunnel retabli avec succes sur: {new_url}")
        else:
            # Refresh shortcut just in case
            if not os.path.exists(HTML_SHORTCUT):
                update_desktop_shortcut(url)

        time.sleep(20)

if __name__ == "__main__":
    main()

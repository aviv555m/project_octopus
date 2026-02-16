from __future__ import annotations

from flask import Flask, render_template_string

app = Flask(__name__)

HTML = """
<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>Squid AI Web v3 (Server Gate)</title>
  <style>
    :root {
      --bg: #0b1020;
      --panel: #0f172a;
      --soft: #1f2937;
      --text: #e2e8f0;
      --muted: #94a3b8;
      --accent: #38bdf8;
      --danger: #f87171;
    }
    body {
      margin: 0;
      background: radial-gradient(circle at top, #1e293b, #0a0f1f 60%);
      color: var(--text);
      font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 16px;
      box-sizing: border-box;
    }
    .app {
      width: min(980px, 100%);
      background: color-mix(in srgb, var(--panel) 92%, black 8%);
      border: 1px solid #334155;
      border-radius: 16px;
      overflow: hidden;
      box-shadow: 0 24px 60px rgba(0, 0, 0, 0.45);
    }
    .header {
      padding: 14px 16px;
      border-bottom: 1px solid #334155;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      background: linear-gradient(90deg, #0f172a, #1e293b);
    }
    .title { font-weight: 700; letter-spacing: 0.2px; }
    .sub { color: var(--muted); font-size: 13px; }
    .settings {
      display: grid;
      grid-template-columns: 1fr 1fr auto;
      gap: 8px;
      padding: 12px 16px;
      border-bottom: 1px solid #334155;
      background: #0b1226;
      align-items: center;
    }
    .settings input {
      width: 100%;
      padding: 8px 10px;
      border-radius: 8px;
      border: 1px solid #334155;
      background: #020617;
      color: var(--text);
      outline: none;
      font-size: 13px;
    }
    .settings button {
      border: 0;
      border-radius: 8px;
      background: var(--accent);
      color: #082f49;
      font-weight: 700;
      padding: 8px 14px;
      cursor: pointer;
      font-size: 13px;
    }
    .hint {
      grid-column: 1 / -1;
      color: var(--muted);
      font-size: 12px;
    }
    .chat {
      height: min(62vh, 620px);
      overflow-y: auto;
      padding: 16px;
      background: linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0));
    }
    .msg {
      max-width: 82%;
      margin: 0 0 12px;
      padding: 10px 12px;
      border-radius: 12px;
      line-height: 1.4;
      white-space: pre-wrap;
      word-wrap: break-word;
      border: 1px solid transparent;
    }
    .user { margin-left: auto; background: #0ea5e9; border-color: #38bdf8; color: #082f49; }
    .bot { background: var(--soft); border-color: #334155; }
    .meta { color: var(--muted); font-size: 12px; margin-top: 4px; }
    .composer {
      border-top: 1px solid #334155;
      padding: 12px;
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 10px;
      background: #0b1226;
    }
    textarea {
      resize: none;
      min-height: 72px;
      max-height: 220px;
      padding: 10px;
      border-radius: 10px;
      border: 1px solid #334155;
      background: #020617;
      color: var(--text);
      outline: none;
    }
    button {
      border: 0;
      border-radius: 10px;
      background: var(--accent);
      color: #082f49;
      font-weight: 700;
      padding: 0 18px;
      cursor: pointer;
    }
    button:disabled { opacity: 0.6; cursor: default; }

    @media (max-width: 720px) {
      body { padding: 8px; }
      .app { border-radius: 12px; }
      .header { flex-direction: column; align-items: flex-start; }
      .settings { grid-template-columns: 1fr; }
      .chat { height: 60vh; padding: 12px; }
      .msg { max-width: 100%; }
      .composer { grid-template-columns: 1fr; }
      textarea { min-height: 84px; }
      button { width: 100%; height: 44px; }
    }
  </style>
</head>
<body>
  <div class=\"app\">
    <div class=\"header\">
      <div>
        <div class=\"title\">Squid AI Web v3</div>
        <div class=\"sub\">Browser UI that calls your Squid API server</div>
      </div>
    </div>
    <div class=\"settings\">
      <input id=\"baseUrl\" placeholder=\"Server base URL (e.g. https://your-domain.com)\" />
      <input id=\"apiKey\" placeholder=\"API key (X-API-Key)\" />
      <button id=\"saveBtn\">Save</button>
      <div class=\"hint\">This page calls `POST /chat` on your server. The server must allow CORS for this origin.</div>
    </div>
    <div id=\"chat\" class=\"chat\"></div>
    <div class=\"composer\">
      <textarea id=\"input\" placeholder=\"Type a message...\"></textarea>
      <button id=\"send\">Send</button>
    </div>
  </div>

  <script>
    const chatEl = document.getElementById('chat');
    const inputEl = document.getElementById('input');
    const sendBtn = document.getElementById('send');
    const baseUrlEl = document.getElementById('baseUrl');
    const apiKeyEl = document.getElementById('apiKey');
    const saveBtn = document.getElementById('saveBtn');

    function loadSettings() {
      baseUrlEl.value = localStorage.getItem('squid_base_url') || '';
      apiKeyEl.value = localStorage.getItem('squid_api_key') || '';
    }

    function saveSettings() {
      localStorage.setItem('squid_base_url', baseUrlEl.value.trim());
      localStorage.setItem('squid_api_key', apiKeyEl.value.trim());
    }

    function addMessage(text, who, meta = '') {
      const wrap = document.createElement('div');
      const bubble = document.createElement('div');
      bubble.className = `msg ${who}`;
      bubble.textContent = text;
      wrap.appendChild(bubble);
      if (meta) {
        const m = document.createElement('div');
        m.className = 'meta';
        m.textContent = meta;
        wrap.appendChild(m);
      }
      chatEl.appendChild(wrap);
      chatEl.scrollTop = chatEl.scrollHeight;
    }

    function normalizeBaseUrl(url) {
      return url.replace(/\/+$/, '');
    }

    async function sendMessage() {
      const text = inputEl.value.trim();
      if (!text) return;

      const baseUrl = normalizeBaseUrl(baseUrlEl.value.trim());
      const apiKey = apiKeyEl.value.trim();
      if (!baseUrl) {
        addMessage('Missing server URL. Enter it above first.', 'bot', 'config');
        return;
      }

      addMessage(text, 'user');
      inputEl.value = '';
      sendBtn.disabled = true;

      try {
        const res = await fetch(`${baseUrl}/chat`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'X-API-Key': apiKey
          },
          body: JSON.stringify({ text })
        });

        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          addMessage(data.detail || data.error || 'Request failed.', 'bot', `error ${res.status}`);
          return;
        }

        let meta = data.route ? `route=${data.route}${data.web_provider ? `, web=${data.web_provider}` : ''}` : '';
        if (Array.isArray(data.sources) && data.sources.length) {
          const hosts = data.sources.map(s => {
            try { return new URL(s).host; } catch { return s; }
          }).slice(0, 3);
          meta = meta ? `${meta}, sources=${hosts.join(', ')}` : `sources=${hosts.join(', ')}`;
        }
        addMessage(data.text || '(empty)', 'bot', meta);
      } catch (err) {
        addMessage(String(err), 'bot', 'network error');
      } finally {
        sendBtn.disabled = false;
      }
    }

    saveBtn.addEventListener('click', () => {
      saveSettings();
      addMessage('Saved server settings.', 'bot', 'system');
    });

    sendBtn.addEventListener('click', sendMessage);
    inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendMessage();
      }
    });

    loadSettings();
  </script>
</body>
</html>
"""


@app.get("/")
def index() -> str:
    return render_template_string(HTML)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5002, debug=False)

from __future__ import annotations

import os
import tempfile
import threading
import time
import uuid
import wave
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Optional

from flask import Flask, jsonify, render_template_string, request, send_from_directory

from squid_ai_v2 import SquidRobotBrain, clean_control_chars, pyttsx3

try:
    import requests  # type: ignore
except Exception:
    requests = None  # type: ignore

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

app = Flask(__name__)

brain = SquidRobotBrain(
    tts_enabled=False,
    web_enabled=True,
)

state_lock = threading.Lock()

TTS_DIR = Path(tempfile.gettempdir()) / "squid_web_tts_v2"
TTS_DIR.mkdir(parents=True, exist_ok=True)

HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Squid AI Web v2</title>
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
      width: min(920px, 100%);
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
    .toggles { display: flex; align-items: center; gap: 12px; font-size: 13px; }
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

    @media (max-width: 640px) {
      body { padding: 8px; }
      .app { border-radius: 12px; }
      .header { flex-direction: column; align-items: flex-start; }
      .chat { height: 60vh; padding: 12px; }
      .msg { max-width: 100%; }
      .composer { grid-template-columns: 1fr; }
      textarea { min-height: 84px; }
      button { width: 100%; height: 44px; }
    }
  </style>
</head>
<body>
  <div class="app">
    <div class="header">
      <div>
        <div class="title">Squid AI Web v2</div>
        <div class="sub">Auto-learning + corrections + Minecraft wiki</div>
      </div>
      <div class="toggles">
        <label><input id="speakToggle" type="checkbox" checked> Speak replies</label>
      </div>
    </div>
    <div id="chat" class="chat"></div>
    <div class="composer">
      <textarea id="input" placeholder="Type a message..."></textarea>
      <button id="send">Send</button>
    </div>
  </div>

  <script>
    const chatEl = document.getElementById('chat');
    const inputEl = document.getElementById('input');
    const sendBtn = document.getElementById('send');
    const speakToggle = document.getElementById('speakToggle');

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

    async function sendMessage() {
      const text = inputEl.value.trim();
      if (!text) return;

      addMessage(text, 'user');
      inputEl.value = '';
      sendBtn.disabled = true;

      try {
        const res = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ message: text, speak: speakToggle.checked })
        });

        const data = await res.json();
        if (!res.ok) {
          addMessage(data.error || 'Request failed.', 'bot', 'error');
          return;
        }

        let meta = data.route ? `route=${data.route}${data.web_provider ? `, web=${data.web_provider}` : ''}` : '';
        if (Array.isArray(data.sources) && data.sources.length) {
          const hosts = data.sources.map(s => {
            try { return new URL(s).host; } catch { return s; }
          }).slice(0, 3);
          meta = meta ? `${meta}, sources=${hosts.join(', ')}` : `sources=${hosts.join(', ')}`;
        }
        addMessage(data.answer || '(empty)', 'bot', meta);

        if (speakToggle.checked) {
          if (data.audio_url) {
            const audio = new Audio(data.audio_url);
            audio.play().catch(() => {});
          } else if ('speechSynthesis' in window && data.answer) {
            const utter = new SpeechSynthesisUtterance(data.answer);
            window.speechSynthesis.cancel();
            window.speechSynthesis.speak(utter);
          }
        }
      } catch (err) {
        addMessage(String(err), 'bot', 'network error');
      } finally {
        sendBtn.disabled = false;
      }
    }

    sendBtn.addEventListener('click', sendMessage);
    inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendMessage();
      }
    });
  </script>
</body>
</html>
"""


def _cleanup_tts_files(max_age_seconds: int = 600) -> None:
    now = time.time()
    for file in TTS_DIR.glob("*.wav"):
        try:
            if now - file.stat().st_mtime > max_age_seconds:
                file.unlink(missing_ok=True)
        except OSError:
            pass


def synthesize_tts_audio(text: str) -> Optional[str]:
    if not text.strip():
        return None

    provider = os.getenv("SQUID_TTS_PROVIDER", "local").strip().lower()
    api_key = os.getenv("SQUID_TTS_API_KEY", "").strip()
    voice = os.getenv("SQUID_TTS_VOICE", "").strip()
    region = os.getenv("SQUID_TTS_AZURE_REGION", "").strip() or os.getenv("SQUID_TTS_REGION", "").strip()
    model = os.getenv("SQUID_TTS_MODEL", "gpt-4o-mini-tts").strip()

    if provider != "local" and api_key and requests is not None:
        name = f"{uuid.uuid4().hex}.wav"
        out = TTS_DIR / name
        try:
            if provider == "openai":
                url = "https://api.openai.com/v1/audio/speech"
                headers = {"Authorization": f"Bearer {api_key}"}
                payload = {"model": model, "voice": voice or "alloy", "input": text, "format": "wav"}
                r = requests.post(url, json=payload, headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    out.write_bytes(r.content)
                    _cleanup_tts_files()
                    return f"/api/tts/{name}"
            if provider == "azure" and region:
                url = f"https://{region}.tts.speech.microsoft.com/cognitiveservices/v1"
                vname = voice or "en-US-JennyNeural"
                ssml = f"<speak version='1.0' xml:lang='en-US'><voice name='{vname}'>{text}</voice></speak>"
                headers = {
                    "Ocp-Apim-Subscription-Key": api_key,
                    "Content-Type": "application/ssml+xml",
                    "X-Microsoft-OutputFormat": "riff-24khz-16bit-mono-pcm",
                }
                r = requests.post(url, data=ssml.encode("utf-8"), headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    out.write_bytes(r.content)
                    _cleanup_tts_files()
                    return f"/api/tts/{name}"
            if provider == "elevenlabs":
                voice_id = voice or "EXAVITQu4vr4xnSDxMaL"
                url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}?output_format=pcm_16000"
                headers = {"xi-api-key": api_key, "Content-Type": "application/json"}
                payload = {"text": text, "model_id": "eleven_multilingual_v2"}
                r = requests.post(url, json=payload, headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    pcm = r.content
                    buf = BytesIO()
                    with wave.open(buf, "wb") as wf:
                        wf.setnchannels(1)
                        wf.setsampwidth(2)
                        wf.setframerate(16000)
                        wf.writeframes(pcm)
                    out.write_bytes(buf.getvalue())
                    _cleanup_tts_files()
                    return f"/api/tts/{name}"
        except Exception:
            pass

    if pyttsx3 is None:
        return None

    # Use a short-lived engine just for web playback
    try:
        engine = pyttsx3.init()
        engine.setProperty("rate", 175)
    except Exception:
        return None

    name = f"{uuid.uuid4().hex}.wav"
    out = TTS_DIR / name

    try:
        engine.save_to_file(text, str(out))
        engine.runAndWait()
    except Exception:
        return None

    _cleanup_tts_files()
    return f"/api/tts/{name}"


def process_message(user_text: str) -> Dict[str, Any]:
    user = clean_control_chars(user_text).strip()
    if not user:
        return {"answer": "", "route": "SYSTEM", "reason": "empty"}

    out = brain.reply(user)

    return {
        "answer": out.get("text", ""),
        "route": out.get("route", "GENERAL"),
        "web_provider": out.get("web_provider", "none"),
        "sources": out.get("sources", []),
    }


@app.get("/")
def index() -> str:
    return render_template_string(HTML)


@app.post("/api/chat")
def api_chat():
    payload = request.get_json(silent=True) or {}
    message = str(payload.get("message", ""))

    with state_lock:
        result = process_message(message)

        audio_url = None
        speak_requested = bool(payload.get("speak", True))
        if speak_requested and result.get("answer"):
            audio_url = synthesize_tts_audio(result["answer"])

    result["audio_url"] = audio_url
    return jsonify(result)


@app.get("/api/tts/<path:filename>")
def tts_file(filename: str):
    safe = Path(filename).name
    return send_from_directory(TTS_DIR, safe, mimetype="audio/wav", as_attachment=False)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=False)

Squid AI Project (v2)
=====================

This repo contains:
- `squid_ai_v2.py`: main brain (memory + web + TTS)
- `squid_server.py`: FastAPI server wrapper
- `test.py`: CLI test client with offline TTS
- `web2.py`: optional web UI (Flask)

Below are instructions from zero to running everything.

Prerequisites
-------------
- Python 3.10+ (3.11 recommended)
- Windows PowerShell (examples below)
- Optional: `ollama` installed and running if you use local models
  - Install from https://ollama.com and run `ollama serve`
  - Pull models used in `squid_ai_v2.py` (examples below)

Setup (from zero)
-----------------
1) Open PowerShell in the project folder.

2) Create and activate a virtual environment:
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

3) Install requirements:
```powershell
pip install -r requirements.txt
```

Note: On Windows, `pywin32` is used for the most reliable TTS. If it fails to install, you can still use `pyttsx3`.

Cloud TTS (optional)
--------------------
Set these env vars to use cloud TTS (higher quality):
- `SQUID_TTS_PROVIDER` = `openai` | `azure` | `elevenlabs`
- `SQUID_TTS_API_KEY` = your API key
- `SQUID_TTS_VOICE` = voice name/id (provider-specific)
- `SQUID_TTS_REGION` = Azure region (only for Azure)
- `SQUID_TTS_MODEL` = OpenAI TTS model (default `gpt-4o-mini-tts`)

4) (Optional) If using Ollama, pull the models used in `squid_ai_v2.py`:
```powershell
ollama pull phi3:mini
ollama pull llama3.2:3b
ollama pull qwen2.5:3b
```

Run the CLI brain directly
--------------------------
```powershell
python .\squid_ai_v2.py
```

Run the API server (FastAPI)
----------------------------
1) Set API key (optional; a default key is also baked in for testing):
```powershell
$env:SQUID_API_KEY="lNp0BGmxQzoH435IDxmQNaIaI57FRNcT7LVF0UgCjRs"
```

2) Start server:
```powershell
uvicorn squid_server:app --host 0.0.0.0 --port 8787
```

3) Health check:
```powershell
curl http://127.0.0.1:8787/health
```

Run the test client (TTS + API)
-------------------------------
```powershell
python .\test.py
```

If TTS only speaks once, try the SAPI engine:
```powershell
python .\test.py --tts-engine sapi
```

Use cloud TTS in the test client:
```powershell
$env:SQUID_TTS_PROVIDER="openai"
$env:SQUID_TTS_API_KEY="YOUR_KEY"
python .\test.py --tts-engine cloud --tts-voice "nova"
```

TTS runtime commands in the test client:
- `/voices`             list available voices
- `/voice <index|word>` switch voice (example: `/voice 2` or `/voice female`)
- `/rate 120`           set speech rate
- `/volume 0.8`         set volume

Web UI (optional)
-----------------
```powershell
python .\web2.py
```
Then open: http://127.0.0.1:5001

Cloud TTS (optional, higher quality)
------------------------------------
Set one provider and API key. The app will use cloud TTS first and fall back to local.

```powershell
# Provider: openai | elevenlabs | azure | local
$env:SQUID_TTS_PROVIDER="openai"
$env:SQUID_TTS_API_KEY="YOUR_API_KEY"
$env:SQUID_TTS_VOICE="alloy"   # provider-specific
$env:SQUID_TTS_MODEL="gpt-4o-mini-tts"  # OpenAI only

# Azure only:
$env:SQUID_TTS_AZURE_REGION="eastus"
```

Test client with cloud TTS:
```powershell
python .\test.py --tts-engine cloud
```

Notes
-----
- Web search uses `ddgs`. If it returns no results, check your network or update `ddgs`.
- If TTS speaks only once, re-run the client or try a different voice driver.

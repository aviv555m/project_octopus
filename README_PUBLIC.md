Squid AI Public (v2)
====================

This is the public/demo version of Squid AI. It uses a separate memory file and does not learn or store user profiles (no names, tone preferences, etc).

What’s Included
---------------
- `squid_ai_public.py`: public brain (separate memory file, profile learning disabled)
- `squid_server_public.py`: FastAPI server wrapper for the public brain
- `squid_memory_public.json`: public memory store
- `web3.py`: server-gated browser UI (calls your API server)

Prerequisites
-------------
- Python 3.10+ (3.11 recommended)
- Optional: `ollama` installed and running if you use local models

Setup
-----
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Run the Public API Server
-------------------------
```powershell
$env:SQUID_API_KEY="YOUR_PUBLIC_KEY"
uvicorn squid_server_public:app --host 0.0.0.0 --port 8788
```

Web UI (Local)
-------------
```powershell
python .\web3.py
```
Then open: http://127.0.0.1:5002

Web UI (GitHub Pages)
---------------------
`web3.py` contains a full HTML page inside the `HTML` string. Copy that HTML into an `index.html` and host it on GitHub Pages.

The UI calls `POST /chat` on your server. Your server must allow CORS for your GitHub Pages domain.

Notes
-----
- Public mode does NOT store names or profile preferences.
- Use a separate API key for the public server.

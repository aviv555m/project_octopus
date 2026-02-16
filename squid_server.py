import os
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel

# Import your robot-ready brain
# Make sure squid_ai.py is in the SAME folder and contains SquidRobotBrain
from squid_ai_v2 import SquidRobotBrain

app = FastAPI(title="Squid AI Server", version="1.1")

DEFAULT_SQUID_API_KEY = "lNp0BGmxQzoH435IDxmQNaIaI57FRNcT7LVF0UgCjRs"
SQUID_API_KEY = os.getenv("SQUID_API_KEY", "").strip() or DEFAULT_SQUID_API_KEY
if not SQUID_API_KEY:
    # Fail fast so you don't run an unprotected server by accident
    raise RuntimeError("SQUID_API_KEY is not set in the environment.")

# Create ONE brain instance (kept in memory)
tts_enabled = os.getenv("SQUID_TTS_ENABLED", "0").strip() in {"1", "true", "TRUE", "yes", "YES"}
tts_voice_id = os.getenv("SQUID_TTS_VOICE_ID", "").strip() or None
try:
    tts_rate = int(os.getenv("SQUID_TTS_RATE", "175").strip())
except Exception:
    tts_rate = 175

brain = SquidRobotBrain(
    tts_enabled=tts_enabled,  # default off; enable via env
    tts_voice_id=tts_voice_id,
    tts_rate=tts_rate,
    web_enabled=True,
)

class ChatIn(BaseModel):
    text: str

class ChatOut(BaseModel):
    text: str
    route: str
    used_web: bool
    web_provider: str
    sources: list[str]

def require_api_key(x_api_key: str | None):
    if not x_api_key or x_api_key.strip() != SQUID_API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/chat", response_model=ChatOut)
def chat(payload: ChatIn, x_api_key: str | None = Header(default=None, alias="X-API-Key")):
    require_api_key(x_api_key)

    user_text = (payload.text or "").strip()
    if not user_text:
        raise HTTPException(status_code=400, detail="text is empty")

    out = brain.reply(user_text)

    # Return only clean fields (no debug)
    return {
        "text": out.get("text", ""),
        "route": out.get("route", "GENERAL"),
        "used_web": bool(out.get("used_web", False)),
        "web_provider": out.get("web_provider", "none"),
        "sources": out.get("sources", []),
    }

@app.post("/chat-in", response_model=ChatOut)
def chat_in(payload: ChatIn, x_api_key: str | None = Header(default=None, alias="X-API-Key")):
    return chat(payload, x_api_key=x_api_key)

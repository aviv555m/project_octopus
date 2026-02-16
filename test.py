# squid_robot_test.py
# Full end-to-end TEST client for your robot:
# - Sends a question to your Squid API server (with SQUID_API_KEY)
# - Prints the answer
# - Speaks ONLY the answer via offline TTS (pyttsx3)
# - Easy voice switching (by index or keyword)
#
# Install on the robot/client machine:
#   pip install requests pyttsx3
#
# PowerShell env (example):
#   $env:SQUID_API_KEY="YOUR_TOKEN_HERE"
#
# Run:
#   python .\squid_robot_test.py
#
# Optional:
#   python .\squid_robot_test.py --list-voices
#   python .\squid_robot_test.py --voice-index 2
#   python .\squid_robot_test.py --voice-like "female"
#   python .\squid_robot_test.py --url "http://127.0.0.1:8000/ask"

from __future__ import annotations

import argparse
import os
import sys
import threading
import tempfile
import wave
from io import BytesIO
from typing import Any, Dict, Optional

import requests

try:
    import pyttsx3  # type: ignore
except Exception:
    pyttsx3 = None  # type: ignore

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

try:
    import winsound  # type: ignore
except Exception:
    winsound = None  # type: ignore


# -----------------------------
# CONFIG (easy to edit)
# -----------------------------
DEFAULT_API_URL = "http://127.0.0.1:8787/chat"  # change if your server is elsewhere
DEFAULT_TTS_RATE = 175
DEFAULT_TTS_VOLUME = 1.0  # 0.0..1.0
DEFAULT_SQUID_API_KEY = "lNp0BGmxQzoH435IDxmQNaIaI57FRNcT7LVF0UgCjRs"


# -----------------------------
# TTS helpers (async so it never blocks)
# -----------------------------
class SapiTTS:
    def __init__(self) -> None:
        if win32com is None:
            raise RuntimeError("pywin32 is not installed. Run: pip install pywin32")
        self._voice = win32com.client.Dispatch("SAPI.SpVoice")
        self._desired_rate: int = DEFAULT_TTS_RATE
        self._desired_volume: float = DEFAULT_TTS_VOLUME
        self._desired_voice_index: Optional[int] = None

    def list_voices(self) -> None:
        voices = self._voice.GetVoices()
        count = voices.Count
        if count == 0:
            print("No voices found.")
            return
        for i in range(count):
            v = voices.Item(i)
            name = v.GetDescription()
            print(f"{i}: {name}")

    def _find_voice_index(self, voice_like: str) -> Optional[int]:
        voices = self._voice.GetVoices()
        key = voice_like.strip().lower()
        for i in range(voices.Count):
            name = voices.Item(i).GetDescription().lower()
            if key in name:
                return i
        return None

    def pick_voice_id(self, voice_index: Optional[int], voice_like: Optional[str]) -> Optional[int]:
        voices = self._voice.GetVoices()
        if voices.Count == 0:
            return None
        if voice_index is not None:
            if 0 <= voice_index < voices.Count:
                return int(voice_index)
            print(f"[warn] voice_index {voice_index} out of range (0..{voices.Count-1}).")
            return None
        if voice_like:
            idx = self._find_voice_index(voice_like)
            if idx is not None:
                return idx
            print(f"[warn] No voice matched '{voice_like}'. Use --list-voices or '/voices' to see options.")
        return None

    def set_voice(self, voice_id: Optional[int]) -> None:
        if voice_id is None:
            return
        voices = self._voice.GetVoices()
        if 0 <= voice_id < voices.Count:
            self._desired_voice_index = voice_id

    def set_rate(self, rate: int) -> None:
        self._desired_rate = int(rate)

    def set_volume(self, volume: float) -> None:
        self._desired_volume = float(volume)

    def speak(self, text: str) -> None:
        if not text.strip():
            return
        threading.Thread(target=self._speak_once, args=(text,), daemon=True).start()

    def _speak_once(self, text: str) -> None:
        try:
            self._voice.Rate = max(-10, min(10, int((self._desired_rate - 180) / 10)))
            self._voice.Volume = int(max(0, min(100, self._desired_volume * 100)))
            if self._desired_voice_index is not None:
                voices = self._voice.GetVoices()
                if 0 <= self._desired_voice_index < voices.Count:
                    self._voice.Voice = voices.Item(self._desired_voice_index)
            self._voice.Speak(text)
        except Exception:
            pass


class AsyncTTS:
    def __init__(self) -> None:
        if pyttsx3 is None:
            raise RuntimeError("pyttsx3 is not installed. Run: pip install pyttsx3")
        self._desired_voice_id: Optional[str] = None
        self._desired_rate: int = DEFAULT_TTS_RATE
        self._desired_volume: float = DEFAULT_TTS_VOLUME

    def list_voices(self) -> None:
        try:
            engine = pyttsx3.init()
        except Exception:
            print("[warn] Failed to init TTS engine for voice listing.")
            return
        voices = engine.getProperty("voices") or []
        if not voices:
            print("No voices found.")
            return
        for i, v in enumerate(voices):
            name = getattr(v, "name", "Unknown")
            vid = getattr(v, "id", "")
            langs = getattr(v, "languages", None)
            gender = getattr(v, "gender", None)
            age = getattr(v, "age", None)
            extra = []
            if langs:
                extra.append(f"langs={langs}")
            if gender:
                extra.append(f"gender={gender}")
            if age:
                extra.append(f"age={age}")
            extra_str = (" | " + ", ".join(extra)) if extra else ""
            print(f"{i}: {name}  (id={vid}){extra_str}")

    def _get_voice_id(self, voice_index: Optional[int], voice_like: Optional[str]) -> Optional[str]:
        try:
            engine = pyttsx3.init()
        except Exception:
            return None
        voices = engine.getProperty("voices") or []
        if not voices:
            return None
        if voice_index is not None:
            if 0 <= voice_index < len(voices):
                return getattr(voices[voice_index], "id", None)
            print(f"[warn] voice_index {voice_index} out of range (0..{len(voices)-1}).")
            return None
        if voice_like:
            key = voice_like.strip().lower()
            for v in voices:
                name = str(getattr(v, "name", "")).lower()
                vid = str(getattr(v, "id", "")).lower()
                langs = str(getattr(v, "languages", "")).lower()
                hay = f"{name} {vid} {langs}"
                if key in hay:
                    return getattr(v, "id", None)
            print(f"[warn] No voice matched '{voice_like}'. Use --list-voices or '/voices' to see options.")
        return None

    def pick_voice_id(self, voice_index: Optional[int], voice_like: Optional[str]) -> Optional[str]:
        return self._get_voice_id(voice_index, voice_like)

    def set_voice(self, voice_id: Optional[str]) -> None:
        self._desired_voice_id = voice_id

    def set_rate(self, rate: int) -> None:
        self._desired_rate = int(rate)

    def set_volume(self, volume: float) -> None:
        self._desired_volume = float(volume)

    def speak(self, text: str) -> None:
        if not text.strip():
            return
        threading.Thread(target=self._speak_once, args=(text,), daemon=True).start()

    def _speak_once(self, text: str) -> None:
        try:
            engine = pyttsx3.init()
            try:
                engine.setProperty("rate", max(80, min(350, int(self._desired_rate))))
                engine.setProperty("volume", max(0.0, min(1.0, float(self._desired_volume))))
                if self._desired_voice_id:
                    engine.setProperty("voice", self._desired_voice_id)
            except Exception:
                pass
            engine.say(text)
            engine.runAndWait()
        except Exception:
            pass


class CloudTTS:
    def __init__(self, provider: str, api_key: str, voice: str = "", region: str = "", model: str = "") -> None:
        if not api_key:
            raise RuntimeError("Cloud TTS requires SQUID_TTS_API_KEY.")
        self.provider = provider.lower()
        self.api_key = api_key
        self.voice = voice
        self.region = region
        self.model = model or "gpt-4o-mini-tts"
        self._desired_rate: int = DEFAULT_TTS_RATE
        self._desired_volume: float = DEFAULT_TTS_VOLUME

    def list_voices(self) -> None:
        print("[info] Cloud TTS voices are provider-specific. Set --tts-voice or SQUID_TTS_VOICE.")

    def pick_voice_id(self, voice_index: Optional[int], voice_like: Optional[str]) -> Optional[str]:
        if voice_like:
            return voice_like
        return None

    def set_voice(self, voice_id: Optional[str]) -> None:
        if voice_id:
            self.voice = voice_id

    def set_rate(self, rate: int) -> None:
        self._desired_rate = int(rate)

    def set_volume(self, volume: float) -> None:
        self._desired_volume = float(volume)

    def speak(self, text: str) -> None:
        if not text.strip():
            return
        threading.Thread(target=self._speak_once, args=(text,), daemon=True).start()

    def _play_wav(self, data: bytes, name: str) -> None:
        if winsound is None:
            return
        try:
            tmp = Path(tempfile.gettempdir()) / name
            tmp.write_bytes(data)
            winsound.PlaySound(str(tmp), winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception:
            pass

    def _speak_once(self, text: str) -> None:
        try:
            if self.provider == "openai":
                url = "https://api.openai.com/v1/audio/speech"
                headers = {"Authorization": f"Bearer {self.api_key}"}
                payload = {"model": self.model, "voice": self.voice or "nova", "input": text, "format": "wav"}
                r = requests.post(url, json=payload, headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    self._play_wav(r.content, "squid_tts_openai.wav")
            elif self.provider == "azure":
                if not self.region:
                    return
                url = f"https://{self.region}.tts.speech.microsoft.com/cognitiveservices/v1"
                voice = self.voice or "en-US-JennyNeural"
                ssml = f"""<speak version='1.0' xml:lang='en-US'><voice name='{voice}'>{text}</voice></speak>"""
                headers = {
                    "Ocp-Apim-Subscription-Key": self.api_key,
                    "Content-Type": "application/ssml+xml",
                    "X-Microsoft-OutputFormat": "riff-24khz-16bit-mono-pcm",
                }
                r = requests.post(url, data=ssml.encode("utf-8"), headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    self._play_wav(r.content, "squid_tts_azure.wav")
            elif self.provider == "elevenlabs":
                voice_id = self.voice or "EXAVITQu4vr4xnSDxMaL"
                url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}?output_format=pcm_16000"
                headers = {"xi-api-key": self.api_key, "Content-Type": "application/json"}
                payload = {"text": text, "model_id": "eleven_multilingual_v2"}
                r = requests.post(url, json=payload, headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    wav_data = self._pcm16_to_wav(r.content, 16000)
                    self._play_wav(wav_data, "squid_tts_11.wav")
        except Exception:
            pass

    def _pcm16_to_wav(self, pcm_data: bytes, sample_rate: int) -> bytes:
        buf = BytesIO()
        with wave.open(buf, "wb") as wf:
            wf.setnchannels(1)
            wf.setsampwidth(2)
            wf.setframerate(sample_rate)
            wf.writeframes(pcm_data)
        return buf.getvalue()


class AsyncTTS:
    def __init__(self) -> None:
        if pyttsx3 is None:
            raise RuntimeError("pyttsx3 is not installed. Run: pip install pyttsx3")
        self._desired_voice_id: Optional[str] = None
        self._desired_rate: int = DEFAULT_TTS_RATE
        self._desired_volume: float = DEFAULT_TTS_VOLUME

    def list_voices(self) -> None:
        try:
            engine = pyttsx3.init()
        except Exception:
            print("[warn] Failed to init TTS engine for voice listing.")
            return
        voices = engine.getProperty("voices") or []
        if not voices:
            print("No voices found.")
            return
        for i, v in enumerate(voices):
            name = getattr(v, "name", "Unknown")
            vid = getattr(v, "id", "")
            langs = getattr(v, "languages", None)
            gender = getattr(v, "gender", None)
            age = getattr(v, "age", None)
            extra = []
            if langs:
                extra.append(f"langs={langs}")
            if gender:
                extra.append(f"gender={gender}")
            if age:
                extra.append(f"age={age}")
            extra_str = (" | " + ", ".join(extra)) if extra else ""
            print(f"{i}: {name}  (id={vid}){extra_str}")

    def _get_voice_id(self, voice_index: Optional[int], voice_like: Optional[str]) -> Optional[str]:
        try:
            engine = pyttsx3.init()
        except Exception:
            return None
        voices = engine.getProperty("voices") or []
        if not voices:
            return None
        if voice_index is not None:
            if 0 <= voice_index < len(voices):
                return getattr(voices[voice_index], "id", None)
            print(f"[warn] voice_index {voice_index} out of range (0..{len(voices)-1}).")
            return None
        if voice_like:
            key = voice_like.strip().lower()
            for v in voices:
                name = str(getattr(v, "name", "")).lower()
                vid = str(getattr(v, "id", "")).lower()
                langs = str(getattr(v, "languages", "")).lower()
                hay = f"{name} {vid} {langs}"
                if key in hay:
                    return getattr(v, "id", None)
            print(f"[warn] No voice matched '{voice_like}'. Use --list-voices or '/voices' to see options.")
        return None

    def pick_voice_id(self, voice_index: Optional[int], voice_like: Optional[str]) -> Optional[str]:
        return self._get_voice_id(voice_index, voice_like)

    def set_voice(self, voice_id: Optional[str]) -> None:
        self._desired_voice_id = voice_id

    def set_rate(self, rate: int) -> None:
        self._desired_rate = int(rate)

    def set_volume(self, volume: float) -> None:
        self._desired_volume = float(volume)

    def speak(self, text: str) -> None:
        if not text.strip():
            return
        threading.Thread(target=self._speak_once, args=(text,), daemon=True).start()

    def _speak_once(self, text: str) -> None:
        try:
            engine = pyttsx3.init()
            try:
                engine.setProperty("rate", max(80, min(350, int(self._desired_rate))))
                engine.setProperty("volume", max(0.0, min(1.0, float(self._desired_volume))))
                if self._desired_voice_id:
                    engine.setProperty("voice", self._desired_voice_id)
            except Exception:
                pass
            engine.say(text)
            engine.runAndWait()
        except Exception:
            pass


class CloudTTS:
    def __init__(self, provider: str, api_key: str, voice: str = "", region: str = "", model: str = "") -> None:
        self.provider = provider.lower()
        self.api_key = api_key
        self.voice = voice
        self.region = region
        self.model = model or "gpt-4o-mini-tts"

    def list_voices(self) -> None:
        print("[info] Cloud TTS voice listing is provider-specific. Use provider dashboard or docs.")

    def pick_voice_id(self, voice_index: Optional[int], voice_like: Optional[str]) -> Optional[str]:
        if voice_like:
            return voice_like
        return None

    def set_voice(self, voice_id: Optional[str]) -> None:
        if voice_id:
            self.voice = voice_id

    def set_rate(self, rate: int) -> None:
        pass

    def set_volume(self, volume: float) -> None:
        pass

    def speak(self, text: str) -> None:
        if not text.strip():
            return
        threading.Thread(target=self._speak_once, args=(text,), daemon=True).start()

    def _speak_once(self, text: str) -> None:
        try:
            import requests
            if self.provider == "openai":
                url = "https://api.openai.com/v1/audio/speech"
                headers = {"Authorization": f"Bearer {self.api_key}"}
                payload = {"model": self.model, "voice": self.voice or "alloy", "input": text, "format": "wav"}
                r = requests.post(url, json=payload, headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    self._play_wav_bytes(r.content)
            elif self.provider == "azure":
                if not self.region:
                    return
                url = f"https://{self.region}.tts.speech.microsoft.com/cognitiveservices/v1"
                voice = self.voice or "en-US-JennyNeural"
                ssml = f"<speak version='1.0' xml:lang='en-US'><voice name='{voice}'>{text}</voice></speak>"
                headers = {
                    "Ocp-Apim-Subscription-Key": self.api_key,
                    "Content-Type": "application/ssml+xml",
                    "X-Microsoft-OutputFormat": "riff-24khz-16bit-mono-pcm",
                }
                r = requests.post(url, data=ssml.encode("utf-8"), headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    self._play_wav_bytes(r.content)
            elif self.provider == "elevenlabs":
                voice_id = self.voice or "EXAVITQu4vr4xnSDxMaL"
                url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}?output_format=pcm_16000"
                headers = {
                    "xi-api-key": self.api_key,
                    "Content-Type": "application/json",
                }
                payload = {"text": text, "model_id": "eleven_multilingual_v2"}
                r = requests.post(url, json=payload, headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    self._play_pcm_as_wav(r.content, 16000)
        except Exception:
            pass

    def _play_wav_bytes(self, wav_bytes: bytes) -> None:
        try:
            import winsound  # type: ignore
        except Exception:
            return
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as f:
                f.write(wav_bytes)
                path = f.name
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_SYNC)
        except Exception:
            pass

    def _play_pcm_as_wav(self, pcm: bytes, sample_rate: int) -> None:
        try:
            buf = BytesIO()
            with wave.open(buf, "wb") as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(sample_rate)
                wf.writeframes(pcm)
            self._play_wav_bytes(buf.getvalue())
        except Exception:
            pass


# -----------------------------
# API client
# -----------------------------
def get_api_key() -> str:
    key = os.getenv("SQUID_API_KEY", "").strip() or DEFAULT_SQUID_API_KEY
    if not key:
        raise RuntimeError(
            "SQUID_API_KEY is not set.\n"
            "PowerShell example:\n"
            '  $env:SQUID_API_KEY="your_token_here"\n'
        )
    return key


def ask_squid(api_url: str, api_key: str, question: str, timeout: float = 25.0) -> Dict[str, Any]:
    """
    Expected server response (example):
      {"text":"...", "route":"GENERAL", "used_web":true, "sources":[...]}
    """
    headers = {
        "X-API-Key": api_key,
        "Content-Type": "application/json",
        "User-Agent": "SquidRobotTest/1.0",
    }
    payload = {"text": question}  # <-- if your server uses {"prompt": ...} change this line
    r = requests.post(api_url, json=payload, headers=headers, timeout=timeout)
    if r.status_code != 200:
        body_hint = (r.text or "")[:300].replace("\n", " ")
        raise RuntimeError(f"API error HTTP {r.status_code}. body_hint={body_hint}")
    data = r.json()
    if not isinstance(data, dict) or "text" not in data:
        raise RuntimeError(f"Unexpected API response shape: {data}")
    return data


# -----------------------------
# Main
# -----------------------------
def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--url", default=DEFAULT_API_URL, help="Squid API URL (POST endpoint)")
    p.add_argument("--list-voices", action="store_true", help="List available TTS voices and exit")
    p.add_argument("--voice-index", type=int, default=None, help="Pick voice by index (see --list-voices)")
    p.add_argument("--voice-like", type=str, default=None, help="Pick voice by keyword (e.g., 'female', 'david', 'hebrew')")
    p.add_argument("--tts-voice", type=str, default=None, help="Cloud TTS voice name/id (overrides env)")
    p.add_argument("--rate", type=int, default=DEFAULT_TTS_RATE, help="TTS rate (80..350)")
    p.add_argument("--volume", type=float, default=DEFAULT_TTS_VOLUME, help="TTS volume (0.0..1.0)")
    p.add_argument("--tts-engine", type=str, default="auto", choices=["auto", "cloud", "sapi", "pyttsx3"], help="TTS engine")
    args = p.parse_args()

    # Setup TTS
    tts = None
    provider = os.getenv("SQUID_TTS_PROVIDER", "").strip().lower()
    api_key = os.getenv("SQUID_TTS_API_KEY", "").strip()
    cloud_voice = (args.tts_voice or os.getenv("SQUID_TTS_VOICE", "")).strip()
    cloud_region = os.getenv("SQUID_TTS_AZURE_REGION", "").strip() or os.getenv("SQUID_TTS_REGION", "").strip()
    cloud_model = os.getenv("SQUID_TTS_MODEL", "").strip()

    if args.tts_engine in ("auto", "cloud") and provider and api_key:
        tts = CloudTTS(provider, api_key, voice=cloud_voice, region=cloud_region, model=cloud_model)
    elif args.tts_engine in ("auto", "sapi") and win32com is not None:
        tts = SapiTTS()
    elif args.tts_engine in ("auto", "pyttsx3") and pyttsx3 is not None:
        tts = AsyncTTS()
    else:
        print("[warn] No TTS engine available. Install pywin32 or pyttsx3, or set cloud env vars.")

    if tts is not None:
        tts.set_rate(int(args.rate))
        tts.set_volume(float(args.volume))

        if args.list_voices:
            tts.list_voices()
            return 0

        voice_id = tts.pick_voice_id(args.voice_index, args.voice_like)
        tts.set_voice(voice_id)

    # API key
    try:
        api_key = get_api_key()
    except Exception as e:
        print(str(e))
        return 2

    print("Squid Robot Test Client")
    print(f"API: {args.url}")
    print("Type 'exit' to quit.\n")
    print("TTS commands: /voices, /voice <index|keyword>, /rate <80-350>, /volume <0.0-1.0>\n")

    while True:
        try:
            q = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if not q:
            continue
        if q.lower() in ("exit", "quit"):
            break
        if q.startswith("/"):
            if tts is None:
                print("[warn] TTS disabled.")
                continue
            parts = q.strip().split(maxsplit=1)
            cmd = parts[0].lower()
            arg = parts[1] if len(parts) > 1 else ""
            if cmd == "/voices":
                tts.list_voices()
                continue
            if cmd == "/voice":
                if arg.isdigit():
                    vid = tts.pick_voice_id(int(arg), None)
                else:
                    vid = tts.pick_voice_id(None, arg)
                tts.set_voice(vid)
                continue
            if cmd == "/rate":
                if arg:
                    tts.set_rate(int(arg))
                continue
            if cmd == "/volume":
                if arg:
                    tts.set_volume(float(arg))
                continue
            print("[warn] Unknown command. Use /voices, /voice, /rate, /volume.")
            continue

        try:
            out = ask_squid(args.url, api_key, q)
        except Exception as e:
            print(f"\n[API ERROR] {e}\n")
            # Speak errors? usually no.
            continue

        answer = str(out.get("text", "")).strip()
        print(f"\nSquid: {answer}\n")

        # Speak ONLY the final answer
        if tts is not None and answer:
            try:
                tts.speak(answer)
            except Exception as e:
                print(f"[TTS warn] {e}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

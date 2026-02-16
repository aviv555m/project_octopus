"""
Squid AI v9.1 — ROBOT-INTEGRATION READY (no slash-commands)

"""

from __future__ import annotations

import ast
import json
import math
import os
import re
import tempfile
import threading
import time
import wave
from io import BytesIO
from dataclasses import dataclass, field
from typing import Any, Dict, List, Literal, Optional, Tuple
from urllib.parse import urlparse

import ollama

# -----------------------------
# Optional TTS (offline)
# -----------------------------
try:
    import pyttsx3  # type: ignore
except Exception:
    pyttsx3 = None  # type: ignore

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

# -----------------------------
# Web search (DDGS + optional SerpApi fallback)
# -----------------------------
try:
    from ddgs import DDGS  # type: ignore
except Exception:
    DDGS = None  # type: ignore

try:
    import requests  # type: ignore
except Exception:
    requests = None  # type: ignore


Route = Literal["THERAPY", "GENERAL", "CREATIVE"]
MEMORY_FILE = "squid_memory.json"

# -----------------------------
# Web config
# -----------------------------
WEB_TIMEOUT_SEC = 8.0
WEB_MAX_SNIPPETS = 6
WEB_MAX_CONTEXT_CHARS = 2600
WEB_MAX_ANSWER_CHARS = 700
WEB_RETRY_ATTEMPTS = 2  # if answer doesn't match question, retry search up to N times

# Block sources you don't want (you said “not wiki maybe”)
# You can edit this list for your robot.
BLOCK_DOMAINS = {
    "m.youtube.com",
    "youtube.com",
    "www.youtube.com",
}
LOW_QUALITY_DOMAINS = {
    "pinterest.com",
    "tr.pinterest.com",
    "grokipedia.com",
}

# If everything is blocked and we get no results, we can optionally allow Wikipedia as last resort:
ALLOW_WIKIPEDIA_LAST_RESORT = True

SERPAPI_ENDPOINT = "https://serpapi.com/search"
SERPAPI_API_KEY_DEFAULT = ""  # optional baked-in key; env var overrides: SERPAPI_API_KEY


# -----------------------------
# Prompts
# -----------------------------

THERAPY_SYSTEM = """You are Squid Companion in THERAPY mode.
Style: warm, calm, reflective, supportive.
Tone: friendly, gentle, a little kid/pet-like (female voice).

Structure (follow this order):
1) Reflect the emotion briefly.
2) Validate the feeling (1 sentence).
3) Ask ONE gentle open-ended question.
4) Offer ONE small coping step only if appropriate (optional).

Hard rules:
- Talk in a casual friendly tone, but do NOT be overly casual or use slang.
- Do NOT claim you are a licensed therapist or doctor.
- Do NOT provide medical/legal instructions.
- If user mentions self-harm or suicide: respond with a brief caring safety message,
  encourage reaching local emergency services and trusted people, ask if they're in immediate danger.
- Keep responses short unless the user asks for more.
"""

GENERAL_SYSTEM = """You are Squid Companion in GENERAL mode.
Style: friendly, playful, gentle, and a little kid/pet-like (female voice).

Hard rules:
- Talk in a friendly way not like an AI, more like a helpful pet/friend.
- Use simple, warm wording. Keep it short.
- Do NOT invent facts about real people/brands/shows.
- Answer exactly what the user asked. Do not add extra scope unless requested.
- Prefer concise, actionable steps.
- If user prefers short answers, keep answers under 3 sentences unless they ask for more.

Formatting rules:
- Avoid LaTeX. Use plain text.
- If user requests "in N sentences" or "N bullet points", follow it.
"""

CREATIVE_SYSTEM = """You are Squid Companion in CREATIVE mode (free talk).
You can tell stories, improv, roleplay (SFW), jokes, worldbuilding.
Tone: playful, gentle, a little kid/pet-like (female voice).

Hard rules:
- Talk in a friendly way not like an AI, more like a fun pet/friend.
- Keep it SFW (no explicit sexual content).
- Don't be cruel or hateful.
- If the user asks for factual advice mid-story, suggest switching to GENERAL mode.
- If the user prefers short answers, keep it short and avoid long paragraphs.

Formatting rules:
- Avoid LaTeX. Use plain text.
- If user requests "in N sentences" or "N bullet points", follow it.
"""

# Dream SMP: dedicated system + starter knowledge (kept general / not a huge lore dump)
DREAM_SMP_SYSTEM = """You are Squid Companion in DREAM SMP mode.
You help answer questions about Dream SMP (DSMP): what it is, who was involved, and high-level lore arcs.

Hard rules:
- Be accurate. If you aren't sure, say so and propose a web-check.
- Keep answers focused on what the user asked (no extra dumping).
- Avoid fanfiction; stick to widely-known summaries.
- Prefer short answers unless the user asks for more.

Starter knowledge (safe, general):
- Dream SMP is a multiplayer Minecraft server that became popular for roleplay/story arcs and creator interactions.
- It involved many Minecraft creators/streamers; membership changed over time.
- "Lore" refers to story events/character arcs created during streams/videos, not real-life facts.
"""

ROUTER_SYSTEM = """You are a router. Choose ONE route: THERAPY, GENERAL, or CREATIVE.

Return ONLY a single-line JSON object and NOTHING ELSE.

Schema:
{"route":"THERAPY|GENERAL|CREATIVE","confidence":0.0,"reason":"short reason"}

Now classify the user's message.
"""

WEB_DECIDER_SYSTEM = """You decide whether we should use web search for the user's question.

Return ONLY a single-line JSON object:
{"use_web":true|false,"query":"search query or empty","reason":"short"}

Rules:
- use_web=true for: "who is", "what is", "tell me about", current facts, public figures, brands, events, dates.
- use_web=false for: small talk (e.g., "how are you", "what's up", greetings), personal questions about the assistant, casual chat, feelings, creative requests, pure opinions, coding help, general explanations.
- If ambiguous, prefer use_web=false and let the assistant ask a short follow-up.
- query should be short, like the main entity/topic.
"""

# Important: refined answers must be clean and NOT include sources in the final output
WEB_REFINE_SYSTEM = """You are Squid Companion using WEB RESULTS.
You must answer ONLY using the provided snippets/links. Do not guess.

Output rules:
- Return ONLY the answer text (no "source:" links, no citations).
- Make it directly answer the user's question.
- Keep it short and clean (2-4 sentences unless user asked otherwise).
- Do NOT add unrelated facts the user did not ask for.
- If the snippets are weak or conflicting, say you're not fully sure and ask ONE short follow-up question.
"""

# Validate relevance and propose better query if needed
WEB_VALIDATE_SYSTEM = """You are a checker.
Given a user question and a draft answer, decide if the answer actually addresses the question.

Return ONLY JSON:
{"ok":true|false,"why":"short","better_query":"short improved search query or empty"}

Rules:
- ok=true only if the answer clearly answers what the user asked.
- If ok=false, suggest a better_query that is short and more specific.
- better_query should not include long instructions; just the refined query.
"""

BEST_EFFORT_SYSTEM = """You are Squid Companion using best-effort mode.
You might not have reliable web results. Provide a short, helpful answer anyway.
If you are unsure, add a brief uncertainty tag like "(Not fully sure.)".
Keep it to 2-3 sentences.
"""

BEST_EFFORT_SYSTEM = """You are Squid Companion.
Give a best-effort answer even if information is incomplete.

Rules:
- Keep it short (1-3 sentences).
- If you are not fully sure, add "(Not fully sure.)" at the end.
- Do not mention web results or sources.
"""


# -----------------------------
# Safety triggers
# -----------------------------
SELF_HARM_PAT = re.compile(
    r"\b("
    r"suicid(e|al)|kill myself|self[- ]?harm|hurt myself|want to die|end it all|"
    r"don't want to live|dont want to live|don't want to be alive|dont want to be alive|"
    r"can't go on|cant go on|no reason to live|wish i were dead|"
    r"i don't think i want to live|i dont think i want to live|"
    r"i don’t think i want to live|i don’t want to live|"
    r"i want to hurt myself|i might hurt myself|i think about hurting myself"
    r")\b",
    re.IGNORECASE,
)

THERAPY_HINTS = [
    r"\b(i feel|i'm feeling|im feeling|sad|depressed|anxious|anxiety|panic|stress|stressed|overwhelmed)\b",
    r"\b(lonely|heartbroken|breakup|relationship|no one likes me|nobody likes me|rejected)\b",
    r"\b(therapy|therapist|mental health|trauma)\b",
]

CREATIVE_HINTS = [
    r"\b(tell me a story|make up a story|roleplay|rp\b|pretend|imagine|once upon a time)\b",
    r"\b(write a|write me a|story about|character)\b",
    r"\b(joke|funny|comedy|skit)\b",
]

GREETINGS_PAT = re.compile(r"^(hi|hello|hey|yo|shalom|sup|what's up|whats up)\.?\s*$", re.IGNORECASE)
SMALL_TALK_PAT = re.compile(
    r"^(hi|hello|hey|yo|shalom|sup|what's up|whats up|how are you|how's it going|hows it going|"
    r"how are u|how r you|you okay|you ok)\??$",
    re.IGNORECASE,
)
ASSISTANT_PERSONAL_PAT = re.compile(
    r"\b(your name|what('?s| is) your name|who are you|are you real|are you a (boy|girl|female|male)|"
    r"what are you|tell me about yourself)\b",
    re.IGNORECASE,
)

# Dream SMP topic trigger
DREAM_SMP_PAT = re.compile(
    r"\b(dream\s*smp|dsmp|l'?manberg|manberg|schlatt|tommyinnit|wilbur|technoblade|tubbo|quackity|sapnap|george|ranboo)\b",
    re.IGNORECASE,
)


def safety_message() -> str:
    return (
        "I’m really sorry you’re feeling this way. You don’t have to handle this alone.\n"
        "If you’re in immediate danger or might hurt yourself, please contact local emergency services right now, "
        "or reach out to someone you trust and stay with them.\n"
        "Are you safe right now?"
    )


# -----------------------------
# Utilities
# -----------------------------
def clean_control_chars(s: str) -> str:
    return "".join(
        ch for ch in s
        if ch == "\n" or ch == "\t" or (32 <= ord(ch) <= 126) or (ord(ch) >= 160)
    )


def extract_sentence_constraint(user_text: str) -> Optional[int]:
    m = re.search(r"\bin\s+(\d+)\s+sentences?\b", user_text, re.IGNORECASE)
    if not m:
        return None
    try:
        n = int(m.group(1))
        return n if 1 <= n <= 10 else None
    except Exception:
        return None


def extract_bullet_constraint(user_text: str) -> Optional[int]:
    m = re.search(r"\b(\d+)\s+bullet\s+points?\b", user_text, re.IGNORECASE)
    if not m:
        return None
    try:
        n = int(m.group(1))
        return n if 1 <= n <= 20 else None
    except Exception:
        return None


def split_sentences(text: str) -> List[str]:
    t = text.strip()
    if not t:
        return []
    parts = re.split(r"(?<=[.!?])\s+", t)
    return [p.strip() for p in parts if p.strip()]


def enforce_n_sentences(answer: str, n: int) -> str:
    sents = split_sentences(answer)
    if len(sents) == n:
        return answer.strip()
    if len(sents) > n:
        return " ".join(sents[:n]).strip()
    while len(sents) < n:
        sents.append("That’s the basic idea.")
    return " ".join(sents[:n]).strip()


def enforce_n_bullets(answer: str, n: int) -> str:
    lines = [ln.strip() for ln in answer.strip().splitlines() if ln.strip()]
    bullets = [ln for ln in lines if ln.startswith(("-", "*", "•"))]
    if bullets:
        return "\n".join(bullets[:n]).strip()

    sents = split_sentences(answer)
    if not sents:
        return answer.strip()
    return "\n".join(f"- {s}" for s in sents[:n]).strip()


def _host(url: str) -> str:
    try:
        return urlparse(url).netloc.lower()
    except Exception:
        return ""


def _serpapi_key() -> Optional[str]:
    env = os.getenv("SERPAPI_API_KEY", "").strip()
    if env:
        return env
    baked = (SERPAPI_API_KEY_DEFAULT or "").strip()
    return baked or None


def _strip_sources_like_text(ans: str) -> str:
    """
    Extra cleanup: remove common "(source: ...)" or "Sources:" lines if the model outputs them anyway.
    """
    a = ans.strip()
    # Remove lines starting with "source" or "sources"
    lines = [ln for ln in a.splitlines() if not re.match(r"^\s*(source|sources)\s*[:\-]", ln.strip(), re.I)]
    a = "\n".join(lines).strip()
    # Remove inline "(source: ...)" patterns
    a = re.sub(r"\(\s*source\s*:\s*https?://[^\)]+\)", "", a, flags=re.I).strip()
    return a


# -----------------------------
# Deterministic Math Solver
# -----------------------------
ALLOWED_BINOPS = {
    ast.Add: lambda a, b: a + b,
    ast.Sub: lambda a, b: a - b,
    ast.Mult: lambda a, b: a * b,
    ast.Div: lambda a, b: a / b,
    ast.FloorDiv: lambda a, b: a // b,
    ast.Mod: lambda a, b: a % b,
    ast.Pow: lambda a, b: a**b,
}
ALLOWED_UNARYOPS = {
    ast.UAdd: lambda a: +a,
    ast.USub: lambda a: -a,
}


def _eval_math_node(node: ast.AST) -> float:
    if isinstance(node, ast.Expression):
        return _eval_math_node(node.body)

    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)

    if isinstance(node, ast.Num):
        return float(node.n)

    if isinstance(node, ast.UnaryOp) and type(node.op) in ALLOWED_UNARYOPS:
        return float(ALLOWED_UNARYOPS[type(node.op)](_eval_math_node(node.operand)))

    if isinstance(node, ast.BinOp) and type(node.op) in ALLOWED_BINOPS:
        left = _eval_math_node(node.left)
        right = _eval_math_node(node.right)
        return float(ALLOWED_BINOPS[type(node.op)](left, right))

    if isinstance(node, ast.Call):
        if isinstance(node.func, ast.Name) and node.func.id in ("factorial", "fact"):
            if len(node.args) != 1:
                raise ValueError("factorial takes 1 arg")
            x = _eval_math_node(node.args[0])
            if not float(x).is_integer() or x < 0 or x > 5000:
                raise ValueError("factorial arg must be integer 0..5000")
            return float(math.factorial(int(x)))
        raise ValueError("Only factorial() is allowed")

    raise ValueError("Unsupported expression")


def try_solve_math(user_text: str) -> Optional[str]:
    t = user_text.strip()
    candidate = re.sub(r"(?i)^\s*what is\s+", "", t)
    candidate = re.sub(r"[=?]+", " ", candidate).strip()

    if not re.search(r"\d", candidate):
        return None
    if re.search(r"[^0-9a-zA-Z\+\-\*\/\%\^\(\)\.\!\s]", candidate):
        return None

    candidate = candidate.replace("^", "**")

    def replace_one_factorial(expr: str) -> Tuple[str, bool]:
        if "!" not in expr:
            return expr, False
        for i, ch in enumerate(expr):
            if ch != "!":
                continue
            j = i - 1
            while j >= 0 and expr[j].isspace():
                j -= 1
            if j < 0:
                continue

            if expr[j] == ")":
                depth = 0
                k = j
                while k >= 0:
                    if expr[k] == ")":
                        depth += 1
                    elif expr[k] == "(":
                        depth -= 1
                        if depth == 0:
                            break
                    k -= 1
                if k < 0:
                    continue
                inside = expr[k : j + 1]
                new_expr = expr[:k] + f"factorial{inside}" + expr[i + 1 :]
                return new_expr, True

            k = j
            while k >= 0 and (expr[k].isdigit() or expr[k] == "." or expr[k].isspace()):
                k -= 1
            token = expr[k + 1 : j + 1].strip()
            if not token:
                continue
            new_expr = expr[: k + 1] + f"factorial({token})" + expr[i + 1 :]
            return new_expr, True

        return expr, False

    changed = True
    while changed:
        candidate, changed = replace_one_factorial(candidate)

    if "!" in candidate:
        return None

    if not re.search(r"(\+|\-|\*|\/|\%|\*\*|factorial|\(|\))", candidate):
        return None

    try:
        tree = ast.parse(candidate, mode="eval")
        val = _eval_math_node(tree)
        if abs(val - round(val)) < 1e-12:
            return str(int(round(val)))
        return str(val)
    except Exception:
        return None


# -----------------------------
# Memory / Profile
# -----------------------------
@dataclass
class UserProfile:
    name: Optional[str] = None
    project: Optional[str] = None
    prefers_short_answers: bool = False
    tone: Optional[Literal["casual", "formal"]] = None
    likes_stories: bool = False
    likes_technical_detail: bool = False


@dataclass
class WebCacheEntry:
    ts: float
    key: str
    answer: str
    sources: List[str] = field(default_factory=list)
    topic: Optional[str] = None  # e.g., "dream_smp"


@dataclass
class Memory:
    profile: UserProfile = field(default_factory=UserProfile)

    history: Dict[Route, List[Dict[str, str]]] = field(
        default_factory=lambda: {"THERAPY": [], "GENERAL": [], "CREATIVE": []}
    )

    auto_notes: List[str] = field(default_factory=list)
    autonotes_enabled: bool = True

    # Web answers remembered here
    web_cache: Dict[str, WebCacheEntry] = field(default_factory=dict)
    # Helps match “same thing asked differently”
    web_aliases: Dict[str, str] = field(default_factory=dict)  # alias_key -> canonical_key

    # Dream SMP accumulated facts (short)
    dream_smp_facts: List[str] = field(default_factory=list)

    # Recent context (small, fast memory)
    last_user: Optional[str] = None
    last_assistant: Optional[str] = None
    last_route: Optional[str] = None
    last_entity: Optional[str] = None
    recent_topics: List[str] = field(default_factory=list)

    # Settings
    tts_enabled: bool = False
    tts_voice_id: Optional[str] = None
    tts_rate: int = 175

    max_turns_per_route: int = 10  # pairs
    web_cache_ttl_sec: int = 60 * 60 * 24 * 21  # 21 days

    def load(self, path: str = MEMORY_FILE) -> None:
        if not os.path.exists(path):
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        try:
            self.profile = UserProfile(**(data.get("profile", {}) or {}))
        except Exception:
            self.profile = UserProfile()

        hist = data.get("history", {})
        for k in ["THERAPY", "GENERAL", "CREATIVE"]:
            self.history[k] = hist.get(k, []) or []

        self.auto_notes = data.get("auto_notes", []) or []
        self.autonotes_enabled = bool(data.get("autonotes_enabled", True))

        self.tts_enabled = bool(data.get("tts_enabled", False))
        self.tts_voice_id = data.get("tts_voice_id")
        self.tts_rate = int(data.get("tts_rate", 175))

        self.web_aliases = data.get("web_aliases", {}) or {}
        self.dream_smp_facts = data.get("dream_smp_facts", []) or []

        self.last_user = data.get("last_user")
        self.last_assistant = data.get("last_assistant")
        self.last_route = data.get("last_route")
        self.last_entity = data.get("last_entity")
        self.recent_topics = data.get("recent_topics", []) or []

        self.web_cache = {}
        raw_cache = data.get("web_cache", {}) or {}
        if isinstance(raw_cache, dict):
            for k, v in raw_cache.items():
                if not isinstance(v, dict):
                    continue
                try:
                    self.web_cache[k] = WebCacheEntry(
                        ts=float(v.get("ts", 0.0)),
                        key=str(v.get("key", k)),
                        answer=str(v.get("answer", "")),
                        sources=list(v.get("sources", []) or []),
                        topic=v.get("topic"),
                    )
                except Exception:
                    continue

    def save(self, path: str = MEMORY_FILE) -> None:
        data = {
            "profile": self.profile.__dict__,
            "history": self.history,
            "auto_notes": self.auto_notes,
            "autonotes_enabled": self.autonotes_enabled,
            "tts_enabled": self.tts_enabled,
            "tts_voice_id": self.tts_voice_id,
            "tts_rate": self.tts_rate,
            "web_aliases": self.web_aliases,
            "dream_smp_facts": self.dream_smp_facts,
            "last_user": self.last_user,
            "last_assistant": self.last_assistant,
            "last_route": self.last_route,
            "last_entity": self.last_entity,
            "recent_topics": self.recent_topics,
            "web_cache": {
                k: {
                    "ts": v.ts,
                    "key": v.key,
                    "answer": v.answer,
                    "sources": v.sources,
                    "topic": v.topic,
                }
                for k, v in self.web_cache.items()
            },
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def add_turn(self, route: Route, user: str, assistant: str) -> None:
        self.history[route].append({"role": "user", "content": user})
        self.history[route].append({"role": "assistant", "content": assistant})
        max_msgs = self.max_turns_per_route * 2
        if len(self.history[route]) > max_msgs:
            self.history[route] = self.history[route][-max_msgs:]


def add_auto_note(mem: Memory, note: str) -> None:
    note = note.strip()
    if not note or note in mem.auto_notes:
        return
    mem.auto_notes.append(note)
    if len(mem.auto_notes) > 60:
        mem.auto_notes = mem.auto_notes[-60:]


def update_recent_context(
    mem: Memory,
    user_text: str,
    assistant_text: str,
    route: Route,
    chosen_query: str = "",
) -> None:
    mem.last_user = user_text.strip()[:1000]
    mem.last_assistant = assistant_text.strip()[:1200]
    mem.last_route = route

    ent = extract_entity_like(user_text) or (chosen_query.strip() if chosen_query else "")
    if ent:
        ent_key = normalize_text_key(ent)
        mem.last_entity = ent_key
        mem.recent_topics.append(ent_key)
        if len(mem.recent_topics) > 12:
            mem.recent_topics = mem.recent_topics[-12:]


AMBIGUOUS_FOLLOWUP_PAT = re.compile(
    r"\b(do you know|who)\s+(i'?m|i am)\s+talking about\b|\bwho (is|am) (that|it|he|she|they)\b",
    re.IGNORECASE,
)


def learn_from_user(mem: Memory, text: str) -> None:
    t = text.strip()

    m = re.search(r"\bmy name is\s+([A-Za-z]+)\b", t, re.IGNORECASE)
    if m:
        name = m.group(1).capitalize()
        mem.profile.name = name
        if mem.autonotes_enabled:
            add_auto_note(mem, f"User name is {name}.")

    if re.search(r"\bi prefer short answers\b", t, re.IGNORECASE):
        mem.profile.prefers_short_answers = True
        if mem.autonotes_enabled:
            add_auto_note(mem, "User prefers short answers.")

    if re.search(r"\bplease be (more )?formal\b", t, re.IGNORECASE):
        mem.profile.tone = "formal"
        if mem.autonotes_enabled:
            add_auto_note(mem, "User prefers a formal tone.")

    if re.search(r"\bplease be (more )?casual\b", t, re.IGNORECASE):
        mem.profile.tone = "casual"
        if mem.autonotes_enabled:
            add_auto_note(mem, "User prefers a casual tone.")

    if re.search(r"\btell me a story\b", t, re.IGNORECASE):
        mem.profile.likes_stories = True
        if mem.autonotes_enabled:
            add_auto_note(mem, "User likes stories.")

    if re.search(r"\b(more technical|technical details|go deeper|more detail)\b", t, re.IGNORECASE):
        mem.profile.likes_technical_detail = True
        if mem.autonotes_enabled:
            add_auto_note(mem, "User likes technical detail.")

    if re.search(r"\b(i'?m|i am)\s+building\s+a\s+(squid|octopus)\s+robot\b", t, re.IGNORECASE) or \
       re.search(r"\b(building|making)\b.*\b(squid|octopus)\b", t, re.IGNORECASE):
        mem.profile.project = "squid companion robot"
        if mem.autonotes_enabled:
            add_auto_note(mem, "User is building a squid companion robot.")


# -----------------------------
# Router (rules + LLM fallback)
# -----------------------------
def extract_first_json_object(text: str) -> Optional[str]:
    start = text.find("{")
    if start == -1:
        return None
    depth = 0
    in_str = False
    escape = False
    for i in range(start, len(text)):
        ch = text[i]
        if in_str:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start : i + 1]
    return None


def rule_router(text: str) -> Optional[Route]:
    if SELF_HARM_PAT.search(text):
        return "THERAPY"
    if GREETINGS_PAT.match(text.strip()):
        return "GENERAL"

    tl = text.lower()
    for pat in CREATIVE_HINTS:
        if re.search(pat, tl):
            return "CREATIVE"
    for pat in THERAPY_HINTS:
        if re.search(pat, tl):
            return "THERAPY"
    return None


def llm_router(model: str, text: str) -> Dict[str, Any]:
    try:
        resp = ollama.chat(
            model=model,
            messages=[
                {"role": "system", "content": ROUTER_SYSTEM},
                {"role": "user", "content": text},
            ],
            options={"temperature": 0.0},
            format="json",
        )
        return json.loads(resp["message"]["content"].strip())
    except Exception:
        pass

    resp = ollama.chat(
        model=model,
        messages=[
            {"role": "system", "content": ROUTER_SYSTEM},
            {"role": "user", "content": text},
        ],
        options={"temperature": 0.0},
    )
    blob = extract_first_json_object(resp["message"]["content"].strip()) or ""
    try:
        return json.loads(blob)
    except Exception:
        return {"route": "GENERAL", "confidence": 0.0, "reason": "router_parse_failed"}


def llm_web_decider(model: str, text: str) -> Dict[str, Any]:
    try:
        resp = ollama.chat(
            model=model,
            messages=[
                {"role": "system", "content": WEB_DECIDER_SYSTEM},
                {"role": "user", "content": text},
            ],
            options={"temperature": 0.0},
            format="json",
        )
        return json.loads(resp["message"]["content"].strip())
    except Exception:
        pass

    resp = ollama.chat(
        model=model,
        messages=[
            {"role": "system", "content": WEB_DECIDER_SYSTEM},
            {"role": "user", "content": text},
        ],
        options={"temperature": 0.0},
    )
    blob = extract_first_json_object(resp["message"]["content"].strip()) or ""
    try:
        return json.loads(blob)
    except Exception:
        return {"use_web": False, "query": "", "reason": "web_decider_parse_failed"}


def rule_web_decider(text: str) -> Optional[Dict[str, Any]]:
    t = text.strip().lower()
    if not t:
        return {"use_web": False, "query": "", "reason": "empty"}
    if SMALL_TALK_PAT.match(t):
        return {"use_web": False, "query": "", "reason": "small_talk"}
    if ASSISTANT_PERSONAL_PAT.search(t):
        return {"use_web": False, "query": "", "reason": "assistant_personal"}
    if try_solve_math(text) is not None:
        return {"use_web": False, "query": "", "reason": "math"}
    # obvious factual prompts
    if re.match(r"^(who is|what is|tell me about|when was|where is|what happened|what happened to)\b", t):
        ent = extract_entity_like(text) or t
        return {"use_web": True, "query": ent, "reason": "rule_obvious_factual"}
    return None


def llm_web_validate(model: str, question: str, answer: str) -> Dict[str, Any]:
    msgs = [
        {"role": "system", "content": WEB_VALIDATE_SYSTEM},
        {"role": "user", "content": f"QUESTION:\n{question}\n\nANSWER:\n{answer}"},
    ]
    try:
        resp = ollama.chat(model=model, messages=msgs, options={"temperature": 0.0}, format="json")
        return json.loads(resp["message"]["content"].strip())
    except Exception:
        pass

    resp = ollama.chat(model=model, messages=msgs, options={"temperature": 0.0})
    blob = extract_first_json_object(resp["message"]["content"].strip()) or ""
    try:
        return json.loads(blob)
    except Exception:
        return {"ok": True, "why": "validate_parse_failed_default_ok", "better_query": ""}


# -----------------------------
# Web search implementations
# -----------------------------
def ddgs_search(query: str, max_results: int = WEB_MAX_SNIPPETS, allow_wiki: bool = False) -> List[Dict[str, str]]:
    """
    Returns list of dicts: {"title","url","snippet"}
    """
    if DDGS is None:
        return []

    try:
        try:
            with DDGS(impersonate="random") as ddgs:
                raw = list(ddgs.text(query, max_results=max_results))
        except TypeError:
            with DDGS() as ddgs:
                raw = list(ddgs.text(query, max_results=max_results))
    except Exception:
        return []

    out: List[Dict[str, str]] = []
    for r in raw:
        title = (r.get("title") or "").strip()
        url = (r.get("href") or r.get("url") or "").strip()
        snippet = (r.get("body") or r.get("snippet") or r.get("description") or "").strip()
        if not url:
            continue

        h = _host(url)
        if (not allow_wiki) and (h in BLOCK_DOMAINS):
            continue
        if allow_wiki:
            # still block youtube always
            if h in {"youtube.com", "www.youtube.com", "m.youtube.com"}:
                continue

        if not snippet and title:
            snippet = title
        out.append({"title": title, "url": url, "snippet": snippet})
    return out


def serpapi_search(query: str, api_key: str, gl: str = "us", hl: str = "en") -> List[Dict[str, str]]:
    """
    Optional fallback if DDGS is blocked.
    Returns {"title","url","snippet"} list.
    """
    if requests is None:
        return []

    params = {"engine": "google", "q": query, "api_key": api_key, "hl": hl, "gl": gl}
    try:
        r = requests.get(
            SERPAPI_ENDPOINT,
            params=params,
            timeout=WEB_TIMEOUT_SEC,
            headers={"User-Agent": "SquidAI/1.0"},
        )
        if r.status_code != 200:
            return []
        data = r.json()
    except Exception:
        return []

    out: List[Dict[str, str]] = []
    org = data.get("organic_results")
    if isinstance(org, list):
        for item in org[:WEB_MAX_SNIPPETS]:
            if not isinstance(item, dict):
                continue
            url = (item.get("link") or "").strip()
            if not url:
                continue
            h = _host(url)
            if h in BLOCK_DOMAINS:
                continue
            out.append(
                {
                    "title": (item.get("title") or "").strip(),
                    "url": url,
                    "snippet": (item.get("snippet") or "").strip(),
                }
            )
    return out


def build_web_context(results: List[Dict[str, str]]) -> Tuple[str, List[str]]:
    lines: List[str] = []
    sources: List[str] = []

    for r in results[:WEB_MAX_SNIPPETS]:
        title = r.get("title", "").strip()
        url = r.get("url", "").strip()
        snip = r.get("snippet", "").strip()
        if not snip and title:
            snip = title
        if not url or not snip:
            continue
        sources.append(url)
        title_part = f"{title} — " if title else ""
        lines.append(f"- {title_part}{snip} (from {url})")

    ctx = "\n".join(lines).strip()
    if len(ctx) > WEB_MAX_CONTEXT_CHARS:
        ctx = ctx[:WEB_MAX_CONTEXT_CHARS].rsplit(" ", 1)[0].strip() + "…"

    sources = list(dict.fromkeys(sources))[:3]
    return ctx, sources


# -----------------------------
# TTS wrapper
# -----------------------------
class TTS:
    def __init__(self):
        self.engine = None
        self.available = pyttsx3 is not None
        self._desired_voice_id: Optional[str] = None
        self._desired_rate: int = 175
        self._sapi = None
        self._auto_voice_set = False
        self._cloud_provider = os.getenv("SQUID_TTS_PROVIDER", "local").strip().lower()
        self._cloud_api_key = os.getenv("SQUID_TTS_API_KEY", "").strip()
        self._cloud_voice = os.getenv("SQUID_TTS_VOICE", "").strip()
        self._cloud_region = os.getenv("SQUID_TTS_AZURE_REGION", "").strip() or os.getenv("SQUID_TTS_REGION", "").strip()
        self._cloud_model = os.getenv("SQUID_TTS_MODEL", "gpt-4o-mini-tts").strip()
        if win32com is not None:
            try:
                self._sapi = win32com.client.Dispatch("SAPI.SpVoice")
                self.available = True
            except Exception:
                self._sapi = None

    def ensure(self, voice_id: Optional[str], rate: int) -> bool:
        if not self.available and self._sapi is None:
            return False
        self._desired_voice_id = voice_id
        self._desired_rate = int(rate)
        # Auto-pick a female/kid-like voice if none provided
        if not self._desired_voice_id:
            self._auto_select_female_voice()
        return True

    def _auto_select_female_voice(self) -> None:
        if self._auto_voice_set:
            return
        # Prefer SAPI voices with "Female" or known female voices
        if self._sapi is not None:
            try:
                voices = self._sapi.GetVoices()
                for i in range(voices.Count):
                    v = voices.Item(i)
                    name = v.GetDescription().lower()
                    if "female" in name or "zira" in name or "susan" in name:
                        self._sapi.Voice = v
                        self._auto_voice_set = True
                        return
            except Exception:
                pass
        # pyttsx3 fallback: pick a voice id that looks female
        if pyttsx3 is not None:
            try:
                engine = pyttsx3.init()
                voices = engine.getProperty("voices") or []
                for v in voices:
                    name = str(getattr(v, "name", "")).lower()
                    vid = str(getattr(v, "id", "")).lower()
                    if "female" in name or "female" in vid or "zira" in name or "susan" in name:
                        self._desired_voice_id = getattr(v, "id", None)
                        self._auto_voice_set = True
                        return
            except Exception:
                pass

    def _speak_once(self, text: str) -> None:
        if self._cloud_provider and self._cloud_provider != "local" and self._cloud_api_key and requests is not None:
            if self._speak_cloud(text):
                return
        if self._sapi is not None:
            try:
                self._sapi.Rate = max(-10, min(10, int((self._desired_rate - 180) / 10)))
                self._sapi.Volume = 100
                self._sapi.Speak(text)
                return
            except Exception:
                pass
        if pyttsx3 is None:
            return
        try:
            engine = pyttsx3.init()
            try:
                engine.setProperty("rate", int(self._desired_rate))
                if self._desired_voice_id:
                    engine.setProperty("voice", self._desired_voice_id)
            except Exception:
                pass
            engine.say(text)
            engine.runAndWait()
        except Exception:
            pass

    def _speak_cloud(self, text: str) -> bool:
        if requests is None:
            return False
        try:
            if self._cloud_provider == "openai":
                url = "https://api.openai.com/v1/audio/speech"
                headers = {"Authorization": f"Bearer {self._cloud_api_key}"}
                payload = {"model": self._cloud_model, "voice": self._cloud_voice or "alloy", "input": text, "format": "wav"}
                r = requests.post(url, json=payload, headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    return self._play_wav_bytes(r.content)
            elif self._cloud_provider == "azure":
                if not self._cloud_region:
                    return False
                url = f"https://{self._cloud_region}.tts.speech.microsoft.com/cognitiveservices/v1"
                voice = self._cloud_voice or "en-US-JennyNeural"
                ssml = (
                    f"<speak version='1.0' xml:lang='en-US'>"
                    f"<voice name='{voice}'>{text}</voice></speak>"
                )
                headers = {
                    "Ocp-Apim-Subscription-Key": self._cloud_api_key,
                    "Content-Type": "application/ssml+xml",
                    "X-Microsoft-OutputFormat": "riff-24khz-16bit-mono-pcm",
                }
                r = requests.post(url, data=ssml.encode("utf-8"), headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    return self._play_wav_bytes(r.content)
            elif self._cloud_provider == "elevenlabs":
                voice_id = self._cloud_voice or "EXAVITQu4vr4xnSDxMaL"
                url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}?output_format=pcm_16000"
                headers = {
                    "xi-api-key": self._cloud_api_key,
                    "Content-Type": "application/json",
                }
                payload = {"text": text, "model_id": "eleven_multilingual_v2"}
                r = requests.post(url, json=payload, headers=headers, timeout=20)
                if r.status_code == 200 and r.content:
                    return self._play_pcm_as_wav(r.content, 16000)
        except Exception:
            return False
        return False

    def _play_wav_bytes(self, wav_bytes: bytes) -> bool:
        try:
            import winsound  # type: ignore
        except Exception:
            return False
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as f:
                f.write(wav_bytes)
                path = f.name
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_SYNC)
            return True
        except Exception:
            return False

    def _play_pcm_as_wav(self, pcm: bytes, sample_rate: int) -> bool:
        try:
            buf = BytesIO()
            with wave.open(buf, "wb") as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(sample_rate)
                wf.writeframes(pcm)
            return self._play_wav_bytes(buf.getvalue())
        except Exception:
            return False

    def say(self, text: str) -> None:
        if not self.available:
            return
        if not text.strip():
            return
        threading.Thread(target=self._speak_once, args=(text,), daemon=True).start()


# -----------------------------
# LLM runner
# -----------------------------
def build_messages(system_prompt: str, mem: Memory, route: Route, user_text: str) -> List[Dict[str, str]]:
    msgs: List[Dict[str, str]] = [{"role": "system", "content": system_prompt}]

    prof_lines: List[str] = []
    if mem.profile.name:
        prof_lines.append(f"User name: {mem.profile.name}")
    if mem.profile.project:
        prof_lines.append(f"Project: {mem.profile.project}")
    prof_lines.append(f"User prefers short answers: {'YES' if mem.profile.prefers_short_answers else 'NO'}")
    if mem.profile.tone:
        prof_lines.append(f"User tone preference: {mem.profile.tone}")
    if mem.profile.likes_stories:
        prof_lines.append("User likes stories: YES")
    if mem.profile.likes_technical_detail:
        prof_lines.append("User likes technical detail: YES")

    msgs.append({"role": "system", "content": "Known user profile:\n- " + "\n- ".join(prof_lines)})

    if mem.auto_notes:
        notes = mem.auto_notes[-10:]
        msgs.append({"role": "system", "content": "Auto-notes:\n- " + "\n- ".join(notes)})

    recent_bits: List[str] = []
    if mem.last_entity:
        recent_bits.append(f"Last entity/topic: {mem.last_entity}")
    if mem.recent_topics:
        recent_bits.append("Recent topics: " + ", ".join(mem.recent_topics[-6:]))
    if mem.last_user and mem.last_assistant:
        recent_bits.append(f"Last exchange: user='{mem.last_user[:160]}' assistant='{mem.last_assistant[:200]}'")
    if recent_bits:
        msgs.append({"role": "system", "content": "Recent context:\n- " + "\n- ".join(recent_bits)})

    msgs.extend(mem.history[route])
    msgs.append({"role": "user", "content": user_text})
    return msgs


def ollama_chat(model: str, messages: List[Dict[str, str]], options: Dict[str, Any]) -> str:
    resp = ollama.chat(model=model, messages=messages, options=options)
    return resp["message"]["content"].strip()


def run_agent(model: str, system_prompt: str, mem: Memory, route: Route, user_text: str) -> str:
    messages = build_messages(system_prompt, mem, route, user_text)

    if mem.profile.tone == "formal":
        messages.insert(1, {"role": "system", "content": "Tone: formal, concise, professional."})
    elif mem.profile.tone == "casual":
        messages.insert(1, {"role": "system", "content": "Tone: casual, friendly, simple wording."})

    if mem.profile.prefers_short_answers:
        cap = 150 if route == "CREATIVE" else 130
        options = {"temperature": 0.6 if route != "CREATIVE" else 0.8, "num_predict": cap}
    else:
        options = {"temperature": 0.75, "num_predict": 360}

    return ollama_chat(model=model, messages=messages, options=options)


def run_web_refine(model: str, mem: Memory, user_text: str, web_context: str) -> str:
    msgs = [
        {"role": "system", "content": WEB_REFINE_SYSTEM},
        {"role": "system", "content": f"WEB RESULTS:\n{web_context}"},
        {"role": "user", "content": user_text},
    ]
    if mem.profile.prefers_short_answers:
        options = {"temperature": 0.25, "num_predict": 190}
    else:
        options = {"temperature": 0.3, "num_predict": 360}

    ans = ollama_chat(model=model, messages=msgs, options=options).strip()
    ans = _strip_sources_like_text(ans)

    if len(ans) > WEB_MAX_ANSWER_CHARS:
        ans = ans[:WEB_MAX_ANSWER_CHARS].rsplit(" ", 1)[0].strip() + "…"
    return ans


def run_best_effort(model: str, mem: Memory, user_text: str) -> str:
    msgs = [
        {"role": "system", "content": BEST_EFFORT_SYSTEM},
        {"role": "user", "content": user_text},
    ]
    if mem.profile.prefers_short_answers:
        options = {"temperature": 0.6, "num_predict": 140}
    else:
        options = {"temperature": 0.7, "num_predict": 220}
    ans = ollama_chat(model=model, messages=msgs, options=options).strip()
    ans = _strip_sources_like_text(ans)
    return ans


# -----------------------------
# Cache helpers (the important “remembering” fixes)
# -----------------------------
def normalize_text_key(t: str) -> str:
    t = t.lower().strip()
    t = re.sub(r"\s+", " ", t)
    t = t.strip(" .,!?:;\"'()[]{}")
    return t


def extract_entity_like(user_text: str) -> Optional[str]:
    """
    Heuristic entity extraction for caching:
    - who is X / what is X / tell me about X / do you know X
    """
    t = normalize_text_key(user_text)
    m = re.match(r"^(who is|what is|tell me about|do you know who|do you know)\s+(.+)$", t)
    if not m:
        return None
    ent = m.group(2).strip().rstrip("?").strip()
    ent = re.sub(r"\bis\b$", "", ent).strip()
    # remove leading articles
    ent = re.sub(r"^(a|an|the)\s+", "", ent).strip()
    return ent or None


def extract_quoted_title(user_text: str) -> Optional[str]:
    m = re.search(r"['\"]([^'\"]{2,80})['\"]", user_text)
    if not m:
        return None
    title = m.group(1).strip()
    return title or None


def expand_queries(user_text: str, chosen_query: str = "") -> List[str]:
    t = user_text.strip()
    quoted = extract_quoted_title(t)
    ent = extract_entity_like(t)
    base = chosen_query.strip() or ent or quoted or t

    candidates: List[str] = []
    def add(q: str) -> None:
        q = q.strip().strip("\"'")
        if q and q not in candidates:
            candidates.append(q)

    add(base)
    if quoted:
        add(quoted)
        add(f"{quoted} movie")
        add(f"{quoted} film")
    if ent:
        add(ent)
        add(f"{ent} official site")
        add(f"{ent} overview")
    return candidates[:6]


def rank_results(results: List[Dict[str, str]], query_terms: List[str]) -> List[Dict[str, str]]:
    terms = [w for w in query_terms if w]
    def score(r: Dict[str, str]) -> int:
        title = (r.get("title") or "").lower()
        snippet = (r.get("snippet") or "").lower()
        url = (r.get("url") or "").lower()
        s = 0
        for w in terms:
            if w in title:
                s += 3
            if w in snippet:
                s += 2
            if w in url:
                s += 1
        host = _host(url)
        if host in LOW_QUALITY_DOMAINS:
            s -= 5
        return s

    return sorted(results, key=score, reverse=True)


def canonical_cache_key(user_text: str, chosen_query: str = "") -> str:
    """
    This is the key fix:
    - Prefer entity keys
    - Else prefer chosen_query (from web-decider)
    - Else fallback to normalized question
    """
    ent = extract_entity_like(user_text)
    if ent:
        return f"ent:{normalize_text_key(ent)}"
    if chosen_query.strip():
        return f"topic:{normalize_text_key(chosen_query)}"
    return f"q:{normalize_text_key(user_text)}"


def question_cache_key(user_text: str) -> str:
    return f"q:{normalize_text_key(user_text)}"


def cache_get_web(mem: Memory, key: str) -> Optional[WebCacheEntry]:
    # Follow alias mapping if exists
    if key in mem.web_aliases:
        key = mem.web_aliases[key]

    ent = mem.web_cache.get(key)
    if not ent:
        return None
    if (time.time() - ent.ts) > mem.web_cache_ttl_sec:
        try:
            del mem.web_cache[key]
        except Exception:
            pass
        return None
    return ent


def cache_set_web(mem: Memory, key: str, answer: str, sources: List[str], topic: Optional[str] = None) -> None:
    mem.web_cache[key] = WebCacheEntry(ts=time.time(), key=key, answer=answer, sources=sources, topic=topic)

    # keep cache small
    if len(mem.web_cache) > 140:
        items = sorted(mem.web_cache.items(), key=lambda kv: kv[1].ts)
        for k, _ in items[:35]:
            mem.web_cache.pop(k, None)


def cache_add_alias(mem: Memory, alias_key: str, canonical_key: str) -> None:
    if alias_key == canonical_key:
        return
    mem.web_aliases[alias_key] = canonical_key
    # keep aliases from growing forever
    if len(mem.web_aliases) > 300:
        # drop oldest-ish by random-ish slicing (simple)
        for k in list(mem.web_aliases.keys())[:80]:
            mem.web_aliases.pop(k, None)


def dream_smp_learn(mem: Memory, question: str, answer: str) -> None:
    """
    Store short “facts learned” from web-refined answers about DSMP.
    Keep it short and dedup.
    """
    if not DREAM_SMP_PAT.search(question):
        return
    fact = answer.strip()
    if not fact:
        return
    # Keep only first ~200 chars per fact
    fact = fact[:200].rsplit(" ", 1)[0].strip()
    if not fact:
        return
    if fact in mem.dream_smp_facts:
        return
    mem.dream_smp_facts.append(fact)
    if len(mem.dream_smp_facts) > 80:
        mem.dream_smp_facts = mem.dream_smp_facts[-80:]


# -----------------------------
# Robot Brain
# -----------------------------
class SquidRobotBrain:
    def __init__(
        self,
        memory_path: str = MEMORY_FILE,
        router_model: str = "phi3:mini",
        therapy_model: str = "llama3.2:3b",
        general_model: str = "qwen2.5:3b",
        creative_model: str = "llama3.2:3b",
        web_refine_model: str = "qwen2.5:3b",
        web_validate_model: str = "phi3:mini",
        dream_smp_model: str = "qwen2.5:3b",
        tts_enabled: bool = False,
        tts_voice_id: Optional[str] = None,
        tts_rate: int = 175,
        web_enabled: bool = True,
        allow_wikipedia_last_resort: bool = ALLOW_WIKIPEDIA_LAST_RESORT,
    ):
        self.memory_path = memory_path
        self.mem = Memory()
        self.mem.load(memory_path)

        # override runtime settings
        self.mem.tts_enabled = bool(tts_enabled)
        self.mem.tts_voice_id = tts_voice_id
        self.mem.tts_rate = int(tts_rate)

        self.router_model = router_model
        self.therapy_model = therapy_model
        self.general_model = general_model
        self.creative_model = creative_model

        self.web_refine_model = web_refine_model
        self.web_validate_model = web_validate_model

        self.dream_smp_model = dream_smp_model

        self.web_enabled = bool(web_enabled)
        self.allow_wikipedia_last_resort = bool(allow_wikipedia_last_resort)

        self.tts = TTS()
        if self.mem.tts_enabled and self.tts.available:
            self.tts.ensure(self.mem.tts_voice_id, self.mem.tts_rate)

    def save(self) -> None:
        self.mem.save(self.memory_path)

    def _speak(self, text: str) -> None:
        if not self.mem.tts_enabled:
            return
        if not self.tts.available:
            return
        if self.tts.ensure(self.mem.tts_voice_id, self.mem.tts_rate):
            self.tts.say(text)

    def _decide_route(self, user_text: str) -> Tuple[Route, float, str]:
        decided = rule_router(user_text)
        if decided is not None:
            return decided, 1.0, "rule"

        meta = llm_router(self.router_model, user_text)
        r = meta.get("route", "GENERAL")
        c = float(meta.get("confidence", 0.0))
        rsn = str(meta.get("reason", "router"))
        if r in ("THERAPY", "GENERAL", "CREATIVE") and c >= 0.55:
            return r, c, rsn  # type: ignore
        return "GENERAL", c, f"router_low_conf({rsn})"

    def _ai_wants_web(self, user_text: str, route: Route) -> Tuple[bool, str, str]:
        if not self.web_enabled:
            return False, "", "web_disabled"
        if route != "GENERAL":
            return False, "", "not_general"
        if try_solve_math(user_text) is not None:
            return False, "", "math"

        rule = rule_web_decider(user_text)
        if rule is not None:
            return bool(rule.get("use_web", False)), str(rule.get("query", "")).strip(), str(rule.get("reason", "rule"))

        meta = llm_web_decider(self.router_model, user_text)
        use_web = bool(meta.get("use_web", False))
        query = str(meta.get("query", "")).strip()
        reason = str(meta.get("reason", "decider"))
        # Safety: if model says use_web but query empty, fallback to entity-like guess
        if use_web and not query:
            ent = extract_entity_like(user_text)
            if ent:
                query = ent
        return use_web, query, reason

    def _web_search(self, query: str) -> Tuple[List[Dict[str, str]], str]:
        # 1) DDGS, strict (no wiki)
        results = ddgs_search(query, max_results=WEB_MAX_SNIPPETS, allow_wiki=False)
        if results:
            return results, "ddgs"

        # 2) DDGS, allow wiki as last resort (optional)
        if self.allow_wikipedia_last_resort:
            results2 = ddgs_search(query, max_results=WEB_MAX_SNIPPETS, allow_wiki=True)
            if results2:
                return results2, "ddgs_wiki_last_resort"

        # 3) SerpApi fallback if key exists
        key = _serpapi_key()
        if key:
            results3 = serpapi_search(query, api_key=key)
            if results3:
                return results3, "serpapi"

        return [], "no_results"

    def _run_dream_smp_answer(self, user: str) -> str:
        """
        Dedicated DSMP mode response. Uses built-in DSMP system prompt + learned DSMP facts as extra context.
        """
        # Inject learned DSMP facts (short)
        learned = ""
        if self.mem.dream_smp_facts:
            take = self.mem.dream_smp_facts[-12:]
            learned = "\nLearned DSMP facts (from previous web checks):\n- " + "\n- ".join(take)

        # Build messages
        msgs: List[Dict[str, str]] = [
            {"role": "system", "content": DREAM_SMP_SYSTEM + learned},
        ]
        # Small general profile injection
        prof_lines = []
        if self.mem.profile.name:
            prof_lines.append(f"User name: {self.mem.profile.name}")
        prof_lines.append(f"User prefers short answers: {'YES' if self.mem.profile.prefers_short_answers else 'NO'}")
        if prof_lines:
            msgs.append({"role": "system", "content": "Known user profile:\n- " + "\n- ".join(prof_lines)})

        msgs.extend(self.mem.history["GENERAL"])
        msgs.append({"role": "user", "content": user})

        if self.mem.profile.prefers_short_answers:
            options = {"temperature": 0.55, "num_predict": 180}
        else:
            options = {"temperature": 0.65, "num_predict": 320}

        ans = ollama_chat(self.dream_smp_model, msgs, options)
        return clean_control_chars(ans).strip()

    def _best_effort_answer(self, user: str) -> str:
        ans = run_best_effort(self.general_model, self.mem, user)
        ans = clean_control_chars(ans).strip()
        return ans

    def reply(self, user_text: str) -> Dict[str, Any]:
        user = clean_control_chars(user_text).strip()
        if not user:
            return {"text": "", "route": "GENERAL", "used_web": False, "web_provider": "none", "sources": []}

        learn_from_user(self.mem, user)

        # Safety
        if SELF_HARM_PAT.search(user):
            answer = safety_message()
            self.mem.add_turn("THERAPY", user, answer)
            update_recent_context(self.mem, user, answer, "THERAPY")
            self.save()
            self._speak(answer)
            return {"text": answer, "route": "THERAPY", "used_web": False, "web_provider": "none", "sources": []}

        # Math
        math_ans = try_solve_math(user)
        if math_ans is not None:
            self.mem.add_turn("GENERAL", user, math_ans)
            update_recent_context(self.mem, user, math_ans, "GENERAL")
            self.save()
            self._speak(math_ans)
            return {"text": math_ans, "route": "GENERAL", "used_web": False, "web_provider": "none", "sources": []}

        sent_n = extract_sentence_constraint(user)
        bullet_n = extract_bullet_constraint(user)

        route, conf, route_reason = self._decide_route(user)

        # Fast resolution for ambiguous follow-ups ("who is he", "do you know who I'm talking about")
        if route == "GENERAL" and AMBIGUOUS_FOLLOWUP_PAT.search(user) and self.mem.last_entity:
            guess = self.mem.last_entity
            answer = f"Do you mean {guess}?"
            if sent_n is not None:
                answer = enforce_n_sentences(answer, sent_n)
            if bullet_n is not None:
                answer = enforce_n_bullets(answer, bullet_n)
            self.mem.add_turn("GENERAL", user, answer)
            update_recent_context(self.mem, user, answer, "GENERAL")
            self.save()
            self._speak(answer)
            return {
                "text": answer,
                "route": "GENERAL",
                "used_web": False,
                "web_provider": "none",
                "sources": [],
                "debug": {"route_conf": conf, "route_reason": route_reason, "ambiguous_followup": True},
            }

        # Dream SMP dedicated path (still "GENERAL", but special handling)
        if route == "GENERAL" and DREAM_SMP_PAT.search(user):
            # If user is asking for factual DSMP stuff, we still allow web below,
            # but first attempt DSMP model answer (fast, local).
            dsmp_ans = self._run_dream_smp_answer(user)
            dsmp_ans = _strip_sources_like_text(dsmp_ans)

            if sent_n is not None:
                dsmp_ans = enforce_n_sentences(dsmp_ans, sent_n)
            if bullet_n is not None:
                dsmp_ans = enforce_n_bullets(dsmp_ans, bullet_n)

            # If DSMP answer is uncertain, consider web (AI decides)
            wants_web, query, web_reason = self._ai_wants_web(user, "GENERAL")
            if wants_web:
                # proceed to web section (below) by not returning yet
                pass
            else:
                self.mem.add_turn("GENERAL", user, dsmp_ans)
                update_recent_context(self.mem, user, dsmp_ans, "GENERAL")
                self.save()
                self._speak(dsmp_ans)
                return {
                    "text": dsmp_ans,
                    "route": "GENERAL",
                    "used_web": False,
                    "web_provider": "none",
                    "sources": [],
                    "debug": {"route_conf": conf, "route_reason": route_reason, "dsmp": True},
                }

        # 1) CACHE CHECK (stronger key + alias)
        # Use AI-chosen query if available (for better canonical key)
        use_web, chosen_query, web_reason = self._ai_wants_web(user, route)
        resolved_entity = self.mem.last_entity if AMBIGUOUS_FOLLOWUP_PAT.search(user) else None
        if resolved_entity and not chosen_query:
            chosen_query = resolved_entity

        alias_key = question_cache_key(user)
        canonical_key = (
            f"ent:{normalize_text_key(resolved_entity)}"
            if resolved_entity
            else canonical_cache_key(user, chosen_query)
        )

        # If the alias key differs from canonical key, map it
        # (This helps “ask again differently” hit the same cached answer)
        cache_add_alias(self.mem, alias_key, canonical_key)

        cached = cache_get_web(self.mem, canonical_key)
        if cached and route == "GENERAL":
            answer = cached.answer
            if sent_n is not None:
                answer = enforce_n_sentences(answer, sent_n)
            if bullet_n is not None:
                answer = enforce_n_bullets(answer, bullet_n)

            self.mem.add_turn(route, user, answer)
            update_recent_context(self.mem, user, answer, route, chosen_query=chosen_query)
            self.save()
            self._speak(answer)
            return {
                "text": answer,
                "route": route,
                "used_web": True,
                "web_provider": "cache",
                "sources": cached.sources,
                "debug": {"route_conf": conf, "route_reason": route_reason, "cache_key": canonical_key},
            }

        # 2) WEB (AI decides) + refine + validate + retry
        if route == "GENERAL" and use_web:
            attempt = 0
            candidates = expand_queries(user, chosen_query)
            current_query = candidates[0] if candidates else user
            provider = "none"
            sources: List[str] = []

            while attempt < max(WEB_RETRY_ATTEMPTS, len(candidates)):
                results, provider = self._web_search(current_query)
                if not results:
                    # try next candidate if available
                    if attempt < len(candidates) - 1:
                        attempt += 1
                        current_query = candidates[attempt]
                        continue
                    break

                # rank results before building context
                results = rank_results(results, current_query)
                web_ctx, sources = build_web_context(results)
                if not web_ctx:
                    if attempt < len(candidates) - 1:
                        attempt += 1
                        current_query = candidates[attempt]
                        continue
                    break

                refined = run_web_refine(self.web_refine_model, self.mem, user, web_ctx)
                refined = clean_control_chars(refined).strip()
                refined = _strip_sources_like_text(refined)

                # Enforce constraints
                if sent_n is not None:
                    refined = enforce_n_sentences(refined, sent_n)
                if bullet_n is not None:
                    refined = enforce_n_bullets(refined, bullet_n)

                # Validate relevance to user question
                check = llm_web_validate(self.web_validate_model, user, refined)
                ok = bool(check.get("ok", True))
                better_query = str(check.get("better_query", "")).strip()

                if ok:
                    # Save to cache so next time it doesn't search
                    cache_set_web(self.mem, canonical_key, refined, sources, topic="dream_smp" if DREAM_SMP_PAT.search(user) else None)

                    # Dream SMP: learn from refined answer
                    dream_smp_learn(self.mem, user, refined)

                    self.mem.add_turn("GENERAL", user, refined)
                    update_recent_context(self.mem, user, refined, "GENERAL", chosen_query=current_query)
                    self.save()
                    self._speak(refined)
                    return {
                        "text": refined,
                        "route": "GENERAL",
                        "used_web": True,
                        "web_provider": provider,
                        "sources": sources,
                        "debug": {
                            "route_conf": conf,
                            "route_reason": route_reason,
                            "web_reason": web_reason,
                            "query": current_query,
                            "cache_key": canonical_key,
                            "validate": check,
                            "attempt": attempt + 1,
                        },
                    }

                # Not OK -> retry with better query (if provided), else tighten query lightly
                attempt += 1
                if better_query:
                    if better_query not in candidates:
                        candidates.insert(0, better_query)
                    current_query = better_query
                else:
                    if attempt < len(candidates) - 1:
                        current_query = candidates[attempt]
                    else:
                        # simple tighten: add extra context words from the question
                        current_query = f"{current_query} {user}".strip()[:120]

            # If web failed, always answer best-effort
            best = self._best_effort_answer(user)
            if "(not fully sure" not in best.lower():
                best = f"{best} (Not fully sure.)"
            self.mem.add_turn("GENERAL", user, best)
            update_recent_context(self.mem, user, best, "GENERAL", chosen_query=current_query)
            self.save()
            self._speak(best)
            return {
                "text": best,
                "route": "GENERAL",
                "used_web": True,
                "web_provider": provider,
                "sources": sources,
                "debug": {"route_conf": conf, "route_reason": route_reason, "web_reason": web_reason, "final_query": current_query},
            }

        # 3) Normal LLM response (no web)
        if route == "THERAPY":
            answer = run_agent(self.therapy_model, THERAPY_SYSTEM, self.mem, route, user)
        elif route == "CREATIVE":
            answer = run_agent(self.creative_model, CREATIVE_SYSTEM, self.mem, route, user)
        else:
            # If it's Dream SMP but no web, we can still use DSMP model
            if DREAM_SMP_PAT.search(user):
                answer = self._run_dream_smp_answer(user)
            else:
                answer = run_agent(self.general_model, GENERAL_SYSTEM, self.mem, route, user)

        answer = clean_control_chars(answer).strip()
        answer = _strip_sources_like_text(answer)

        if sent_n is not None:
            answer = enforce_n_sentences(answer, sent_n)
        if bullet_n is not None:
            answer = enforce_n_bullets(answer, bullet_n)

        self.mem.add_turn(route, user, answer)
        update_recent_context(self.mem, user, answer, route)
        self.save()
        self._speak(answer)

        return {
            "text": answer,
            "route": route,
            "used_web": False,
            "web_provider": "none",
            "sources": [],
            "debug": {"route_conf": conf, "route_reason": route_reason},
        }


# -----------------------------
# Minimal test loop (NO commands)
# -----------------------------
if __name__ == "__main__":
    brain = SquidRobotBrain(
        tts_enabled=True,   # set False if you don't want TTS in CLI test
        web_enabled=True,
        tts_rate=175,
    )
    print("Squid AI v9.1 (robot-ready, no commands). Type 'exit' to quit.\n")
    while True:
        u = input("You: ").strip()
        if not u:
            continue
        if u.lower() in ("exit", "quit"):
            brain.save()
            break
        out = brain.reply(u)
        print(f"\nSquid: {out['text']}\n")

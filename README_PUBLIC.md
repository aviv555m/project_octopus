Squid AI Public (OpenRouter)
============================

This public server uses OpenRouter for all LLM calls.

Quick start
-----------
1) Install deps:
   pip install -r requirements.txt

2) Set environment variables:
   OPENROUTER_API_KEY=your_key_here
   SQUID_API_KEY=your_server_api_key

3) Run the server:
   python squid_server_public.py

Optional model overrides
------------------------
Defaults are `openrouter/auto` for all model roles. You can override any of these:
SQUID_ROUTER_MODEL
SQUID_THERAPY_MODEL
SQUID_GENERAL_MODEL
SQUID_CREATIVE_MODEL
SQUID_WEB_REFINE_MODEL
SQUID_WEB_VALIDATE_MODEL
SQUID_DREAM_SMP_MODEL

Optional OpenRouter headers
---------------------------
OPENROUTER_APP_URL
OPENROUTER_APP_TITLE
OPENROUTER_BASE_URL

Notes
-----
- `squid_memory_public.json` is the public memory store.
- `squid_ai_public.py` contains the OpenRouter client and chat logic.

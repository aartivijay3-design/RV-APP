"""Gemini AI client setup, shared by every feature module."""
import os
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

ai_client = OpenAI(
    # .strip() guards against a stray trailing newline/whitespace in the env
    # var — easy to introduce by pasting from Notepad or a dashboard field,
    # and it silently breaks every request: httpx rejects a "Bearer <key>\n"
    # header as invalid before the request ever leaves the machine, which
    # surfaces only as a generic, misleading "Connection error."
    api_key=os.environ["GEMINI_API_KEY"].strip(),
    base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
    # Without an explicit timeout, a stalled Gemini call (rate limit, network
    # hiccup) can hang for the SDK's default of several minutes — long enough
    # that a user re-analysing a document sees the button stuck on
    # "Analysiert…" with no way to tell a slow response from a dead one.
    # Structure/day-prose completions normally finish in well under this.
    timeout=90.0,
    # The SDK already retries 429/5xx with backoff automatically — this just
    # extends the budget beyond its default of 2. Observed directly this
    # session: Gemini returns 503 "currently experiencing high demand" during
    # brief spikes on Google's side — a shared-infrastructure thing that can
    # hit any caller regardless of plan or billing, not a sign of hitting a
    # quota. 2 attempts wasn't always enough to ride one out, and every one
    # of this app's ~10 call sites benefits from the fix living here rather
    # than needing its own retry loop — a single day silently coming back
    # with no text (and no visible error, if a caller happened to swallow
    # the exception) was hard to tell apart from a genuine content problem.
    max_retries=5,
)
AI_MODEL = "gemini-2.5-flash"


def _ai_complete(**kwargs):
    """Wrapper around ai_client.chat.completions.create with Gemini 'thinking'
    disabled. Gemini 2.5's reasoning tokens silently consume part of max_tokens
    before the visible answer is written — with a modest max_tokens budget this
    can burn almost the whole allowance on invisible reasoning and cut the JSON
    off mid-output (finish_reason="length" with completion_tokens far below
    max_tokens, total_tokens far above prompt+completion). None of this app's
    calls need multi-step reasoning, so thinking is always off.
    """
    return ai_client.chat.completions.create(
        extra_body={"extra_body": {"google": {"thinking_config": {"thinking_budget": 0}}}},
        **kwargs,
    )

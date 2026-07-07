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

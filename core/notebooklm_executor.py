# -*- coding: utf-8 -*-
"""
NotebookLM & AI Content Executor (High-Tolerance JSON Engine)
Handles automated NotebookLM queries, multi-model Gemini AI fallback,
and ultra-robust JSON parsing/repair using json_repair and regex heuristics.
"""

import os
import re
import time
import json
import asyncio
import logging
from typing import Dict, Any, Optional
from google import genai
from google.genai import types

try:
    import json_repair
    HAS_JSON_REPAIR = True
except ImportError:
    HAS_JSON_REPAIR = False

from bot.config import GEMINI_API_KEY

logger = logging.getLogger("NotebookLMExecutor")

try:
    from notebooklm_mcp.config import ServerConfig, AuthConfig
    from notebooklm_mcp.client import NotebookLMClient
    HAS_NOTEBOOKLM_MCP = True
except Exception:
    HAS_NOTEBOOKLM_MCP = False

# Robust ordered list of active 2026 Gemini models
CANDIDATE_MODELS = [
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-flash-latest",
    "gemini-flash-lite-latest",
    "gemini-3.1-flash-lite"
]



class NotebookLMExecutor:
    """
    Executes queries via NotebookLM or resilient Gemini AI fallback,
    and parses responses into structured presentation data with 100% JSON tolerance.
    """
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY")
        self.client = genai.Client(api_key=self.api_key) if self.api_key else None

    async def query_notebooklm_automated(self, prompt: str, timeout: int = 25) -> Optional[str]:
        """
        Attempts to query NotebookLM via browser automation session if active.
        """
        if not HAS_NOTEBOOKLM_MCP:
            return None

        profile_dir = os.environ.get("NOTEBOOKLM_PROFILE_DIR", r"c:\Users\user\chrome_profile")
        if not os.path.exists(profile_dir):
            profile_dir = r"c:\Users\user\.notebooklm_chrome_profile"

        cfg = ServerConfig(
            headless=True,
            timeout=timeout,
            auth=AuthConfig(
                profile_dir=profile_dir,
                use_persistent_session=True,
                auto_login=False
            )
        )
        client = None
        try:
            client = NotebookLMClient(cfg)
            await asyncio.wait_for(client.start(), timeout=8)
            is_auth = await asyncio.wait_for(client.authenticate(), timeout=8)
            if is_auth:
                logger.info("Connected to NotebookLM session successfully.")
                await client.send_message(prompt)
                resp = await client.get_response(wait_for_completion=True, max_wait=timeout)
                if resp and len(resp.strip()) > 50:
                    return resp.strip()
        except Exception as e:
            logger.info(f"NotebookLM automation unavailable ({e}). Using AI fallback.")
        finally:
            if client and hasattr(client, 'close'):
                try:
                    await client.close()
                except Exception:
                    pass
        return None

    def query_groq_fallback(self, prompt: str) -> Optional[str]:
        """
        Ultra-fast resilient fallback using Groq API with zero quota bottleneck.
        """
        groq_key = os.getenv("GROQ_API_KEY")
        if not groq_key:
            return None

        try:
            import requests
            groq_models = ["openai/gpt-oss-120b", "openai/gpt-oss-20b", "qwen/qwen3.8-27b"]
            for g_model in groq_models:
                try:
                    logger.info(f"Trying Groq fallback with model: {g_model}")
                    res = requests.post(
                        "https://api.groq.com/openai/v1/chat/completions",
                        headers={
                            "Authorization": f"Bearer {groq_key}",
                            "Content-Type": "application/json"
                        },
                        json={
                            "model": g_model,
                            "messages": [
                                {
                                    "role": "system",
                                    "content": (
                                        "You are a professional academic presentation generator. "
                                        "Strictly output valid, well-structured JSON only without markdown codeblocks or conversational text."
                                    )
                                },
                                {"role": "user", "content": prompt}
                            ],
                            "response_format": {"type": "json_object"},
                            "temperature": 0.25,
                            "max_tokens": 6144
                        },
                        timeout=40
                    )
                    if res.status_code == 200:
                        data = res.json()
                        choices = data.get("choices", [])
                        if choices:
                            content = choices[0].get("message", {}).get("content", "").strip()
                            if content and len(content) > 30:
                                logger.info(f"Successfully generated presentation content via Groq ({g_model})")
                                return content
                    else:
                        logger.warning(f"Groq model {g_model} returned {res.status_code}: {res.text[:100]}")
                except Exception as e_groq:
                    logger.warning(f"Groq model {g_model} failed: {e_groq}")
        except Exception as e:
            logger.warning(f"Groq fallback overall error: {e}")
        return None

    def query_gemini_fallback(self, prompt: str, max_retries: int = 2) -> str:
        """
        Generates structured slide content using multi-model Gemini fallback with automatic retries
        and seamless Groq API fallback when Gemini quotas are hit.
        """
        last_error = None

        # 1. Try Gemini candidate models if client is initialized
        if self.client:
            for model in CANDIDATE_MODELS:
                for attempt in range(max_retries):
                    try:
                        response = self.client.models.generate_content(
                            model=model,
                            contents=prompt,
                            config=types.GenerateContentConfig(
                                temperature=0.3,
                                max_output_tokens=8192,
                                response_mime_type="application/json"
                            )
                        )
                        if response and response.text and len(response.text.strip()) > 20:
                            logger.info(f"Successfully generated content using Gemini model: {model}")
                            return response.text.strip()
                    except Exception as e:
                        err_str = str(e)
                        logger.warning(f"Gemini model {model} (attempt {attempt+1}) failed: {err_str[:120]}")
                        last_error = e
                        if "503" in err_str or "429" in err_str or "UNAVAILABLE" in err_str or "404" in err_str:
                            break

        # 2. Seamless Fast Fallback: Groq Ultra-Fast API (Zero Quota Bottleneck)
        groq_output = self.query_groq_fallback(prompt)
        if groq_output:
            return groq_output

        # 3. If both Gemini and Groq fail, raise an informative error
        raise RuntimeError(
            f"AI serverlarida vaqtinchalik yuklama yuqori ({last_error}). "
            f"Iltimos, yuqoridagi tayyor promptni nusxalab NotebookLM chatiga tashlang va javobini pastdagi maydonga kiriting."
        )

    async def generate_content_auto(self, prompt: str, try_browser: bool = False) -> Dict[str, Any]:
        """
        High-speed presentation content generator.
        Uses direct multi-model AI (Gemini / Groq) for instant response (< 2 seconds),
        or optional browser automation if explicitly requested.
        """
        raw_output = None
        if try_browser:
            try:
                raw_output = await self.query_notebooklm_automated(prompt)
            except Exception:
                raw_output = None

        if not raw_output:
            raw_output = self.query_gemini_fallback(prompt)

        return self.parse_notebooklm_response(raw_output)

    @staticmethod
    def parse_or_repair_json(raw_text: str) -> Dict[str, Any]:
        """Alias for parse_notebooklm_response: Cleans, repairs, and parses raw JSON or Markdown string."""
        return NotebookLMExecutor.parse_notebooklm_response(raw_text)

    @staticmethod
    def parse_notebooklm_response(raw_text: str) -> Dict[str, Any]:
        """
        Cleans, repairs, and parses raw JSON string from NotebookLM or Gemini output.
        Handles markdown fences, unescaped quotes, trailing commas, comments, and partial truncations.
        """
        if not raw_text or not raw_text.strip():
            raise ValueError("Qabul qilingan matn bo'sh!")

        text = raw_text.strip()

        # Check if text is Markdown Outline (e.g. ### 1-slayd:, Action-oriented, etc.)
        is_markdown_outline = (
            "### " in text or
            "-slayd:" in text.lower() or
            "slide " in text.lower() or
            "action-oriented" in text.lower() or
            "tarkibiy tezislar" in text.lower()
        ) and not (text.startswith("{") or text.startswith("["))

        if is_markdown_outline:
            try:
                from core.notebooklm_markdown_parser import NotebookLMMarkdownParser
                parsed_md = NotebookLMMarkdownParser.parse(text)
                if parsed_md.get("slides"):
                    return parsed_md
            except Exception as e:
                logger.warning(f"Markdown parsing attempt failed: {e}")

        # 1. Strip markdown ```json ... ``` blocks
        if "```" in text:
            match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', text)
            if match:
                text = match.group(1).strip()

        # 2. Remove single line comments // ...
        text = re.sub(r'//.*$', '', text, flags=re.MULTILINE)

        # 3. Direct JSON parse attempt
        try:
            data = json.loads(text)
            if isinstance(data, dict):
                return data
        except Exception:
            pass

        # 4. Use json_repair if available
        if HAS_JSON_REPAIR:
            try:
                repaired = json_repair.loads(text)
                if isinstance(repaired, dict) and "slides" in repaired:
                    return repaired
                elif isinstance(repaired, list) and len(repaired) > 0:
                    return {"topic": "Academic Presentation", "slides": repaired}
            except Exception as e:
                logger.warning(f"json_repair attempt failed: {e}")

        # 5. Regex heuristic repair for trailing commas & braces
        text_repaired = re.sub(r',\s*([\]}])', r'\1', text)
        first_brace = text_repaired.find('{')
        last_brace = text_repaired.rfind('}')
        if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
            extracted = text_repaired[first_brace:last_brace + 1]
            try:
                data = json.loads(extracted)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass

        # 6. If JSON is truncated (unclosed array or object), close brackets
        open_braces = text.count('{') - text.count('}')
        open_brackets = text.count('[') - text.count(']')
        patched_text = text.rstrip().rstrip(',')
        if open_brackets > 0:
            patched_text += ']' * open_brackets
        if open_braces > 0:
            patched_text += '}' * open_braces

        try:
            data = json.loads(patched_text)
            if isinstance(data, dict):
                return data
        except Exception:
            pass

        if HAS_JSON_REPAIR:
            try:
                data = json_repair.loads(patched_text)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass

        # Final fallback: Try NotebookLMMarkdownParser on raw text
        try:
            from core.notebooklm_markdown_parser import NotebookLMMarkdownParser
            parsed_md = NotebookLMMarkdownParser.parse(raw_text)
            if parsed_md.get("slides"):
                return parsed_md
        except Exception:
            pass

        raise ValueError(f"NotebookLM javobidan to'g'ri JSON yoki Markdown formatini ajratib bo'lmadi! Matn:\n{text[:200]}...")


# Singleton instance
_executor_instance: Optional[NotebookLMExecutor] = None

def get_notebooklm_executor() -> NotebookLMExecutor:
    global _executor_instance
    if _executor_instance is None:
        _executor_instance = NotebookLMExecutor()
    return _executor_instance

"""
AI-friendly MCP App Template (FastMCP)

Metadata (JSON) — машинно-четим, AI може да обновява:
{
  "app_name": "ai-mcp-app",
  "description": "Template for AI to generate MCP tools. Tools must accept typed args and return JSON-serializable dicts.",
  "version": "1.0",
  "author": "auto-generated",
  "tool_contracts": [
    {
      "name": "get_current_weather",
      "description": "Returns simple current weather summary for a city.",
      "inputs": {"city": "string"},
      "output": {"status": "string", "generated_at": "string", "city": "string", "summary": "string"}
    }
  ]
}
"""

import os
import json
import time
import logging
from typing import Any, Callable, Dict, Optional, Tuple
from datetime import datetime
import requests
from requests.adapters import HTTPAdapter, Retry
from fastmcp import FastMCP

# -------------------------
# App metadata (AI can edit)
# -------------------------
APP_META = {
    "app_name": os.getenv("MCP_APP_NAME", "ai-mcp-app"),
    "description": "Template for AI-driven MCP tool generation",
    "version": "1.0",
    "author": os.getenv("MCP_APP_AUTHOR", "auto-generated"),
    # AI can append tool_contracts here when generating new tools
    "tool_contracts": []
}

# -------------------------
# Configuration (can be overridden by AI)
# -------------------------
CONFIG = {
    "USER_AGENT": os.getenv("USER_AGENT", "ai-mcp-app/1.0"),
    "CACHE_TTL": int(os.getenv("CACHE_TTL", "300")),
    "REQUEST_TIMEOUT": int(os.getenv("REQUEST_TIMEOUT", "10")),
    "RETRY_TOTAL": int(os.getenv("RETRY_TOTAL", "3")),
    "RETRY_BACKOFF": float(os.getenv("RETRY_BACKOFF", "0.5")),
}

# -------------------------
# Init
# -------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(APP_META["app_name"])
mcp = FastMCP(APP_META["app_name"])

# -------------------------
# Simple cache (AI may replace with Redis)
# -------------------------
_cache: Dict[str, Tuple[float, Any]] = {}

def _now_ts() -> float:
    return time.time()

def cache_get(key: str) -> Optional[Any]:
    entry = _cache.get(key)
    if not entry:
        return None
    ts, value = entry
    if _now_ts() - ts > CONFIG["CACHE_TTL"]:
        _cache.pop(key, None)
        return None
    return value

def cache_set(key: str, value: Any) -> None:
    _cache[key] = (_now_ts(), value)

# -------------------------
# HTTP session factory
# -------------------------
def requests_session() -> requests.Session:
    session = requests.Session()
    retries = Retry(
        total=CONFIG["RETRY_TOTAL"],
        backoff_factor=CONFIG["RETRY_BACKOFF"],
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD", "OPTIONS"]
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.headers.update({"User-Agent": CONFIG["USER_AGENT"]})
    return session

# -------------------------
# Utilities for AI
# -------------------------
def now_iso() -> str:
    return datetime.utcnow().isoformat() + "Z"

def jsonify_safe(obj: Any) -> Any:
    """Ensure obj is JSON-serializable (simple fallback for datetimes)."""
    if isinstance(obj, datetime):
        return obj.isoformat()
    try:
        json.dumps(obj)
        return obj
    except TypeError:
        return str(obj)

# -------------------------
# Dynamic tool registration (for AI)
# -------------------------
def register_tool_from_spec(spec: Dict[str, Any], func: Callable) -> None:
    """
    Register a new tool using a spec dict and a callable.
    Spec example:
    {
      "name": "tool_name",
      "description": "desc",
      "inputs": {"param1": "string", "param2": "int"},
      "output_contract": {"status":"string", "result":"object"}
    }
    AI can generate 'spec' and a Python function 'func' to register dynamically.
    """
    name = spec.get("name")
    if not name:
        raise ValueError("Tool spec must include 'name'")
    # attach metadata for discovery
    APP_META["tool_contracts"].append(spec)
    # register the callable as mcp tool
    mcp.tool(func)

# -------------------------
# Example concrete tools (follow contract: return JSON-serializable dict)
# -------------------------
@mcp.tool
def ping() -> Dict[str, Any]:
    """Health check; minimal contract."""
    return {"status": "ok", "time": now_iso()}

@mcp.tool
def get_current_weather(city: str) -> Dict[str, Any]:
    """
    Example tool: returns a small JSON summary.
    Contract (AI generators should follow):
      input: city (str)
      output: {status, generated_at, city, summary}
    """
    if not city or not city.strip():
        return {"status": "error", "message": "city is empty"}
    cache_key = f"current:{city.lower().strip()}"
    cached = cache_get(cache_key)
    if cached:
        return {"status": "ok", "generated_at": now_iso(), "city": city, "summary": cached, "cached": True}

    # simple implementation using a GET to a public endpoint (AI may change)
    sess = requests_session()
    try:
        url = f"https://wttr.in/{requests.utils.quote(city)}?format=3"
        resp = sess.get(url, timeout=CONFIG["REQUEST_TIMEOUT"])
        resp.raise_for_status()
        summary = resp.text.strip()
        cache_set(cache_key, summary)
        return {"status": "ok", "generated_at": now_iso(), "city": city, "summary": summary}
    except requests.RequestException as e:
        logger.exception("Failed to fetch weather for %s", city)
        return {"status": "error", "message": str(e)}

# -------------------------
# Hooks for AI to extend:
# - add tools by generating functions and calling register_tool_from_spec
# - update APP_META/tool_contracts to describe new tools
# - swap out cache or session implementations
# -------------------------
# Example: AI could run:
# spec = {"name":"echo","description":"Echo input","inputs":{"text":"string"},"output_contract":{"echo":"string"}}
# def echo(text: str): return {"echo": text}
# register_tool_from_spec(spec, echo)

# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    logger.info("%s starting (v%s) at %s", APP_META["app_name"], APP_META["version"], now_iso())
    mcp.run()

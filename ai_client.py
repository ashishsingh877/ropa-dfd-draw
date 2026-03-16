"""
ai_client.py  —  Groq API wrapper
Model: llama-3.3-70b-versatile (best for strict JSON)
Free at console.groq.com
"""

import json, re, requests

GROQ_BASE  = "https://api.groq.com/openai/v1/chat/completions"
ALL_MODELS = [
    "llama-3.3-70b-versatile",
    "llama3-70b-8192",
    "llama-3.1-70b-versatile",
    "mixtral-8x7b-32768",
]

def _body(system: str, user: str, max_tokens: int, model: str) -> dict:
    return {
        "model":       model,
        "temperature": 0.2,
        "max_tokens":  max_tokens,
        "messages": [
            {"role": "system",  "content": system},
            {"role": "user",    "content": user},
        ]
    }

def chat(api_key: str, system: str, user: str,
         max_tokens: int = 8000, model: str = "llama-3.3-70b-versatile") -> str:
    """Blocking Groq call with model fallback."""
    models = [model] + [m for m in ALL_MODELS if m != model]
    errors = []
    for m in models:
        try:
            r = requests.post(
                GROQ_BASE,
                headers={"Authorization": f"Bearer {api_key}",
                         "Content-Type": "application/json"},
                json=_body(system, user, max_tokens, m),
                timeout=120
            )
            if r.status_code == 200:
                return r.json()["choices"][0]["message"]["content"]
            if r.status_code == 401:
                raise ValueError("Invalid Groq API key. Get one free at https://console.groq.com")
            try:    msg = r.json().get("error", {}).get("message", r.text[:300])
            except: msg = r.text[:300]
            errors.append(f"[{m}] {r.status_code}: {msg}")
            if r.status_code == 429:
                continue   # rate limit — try next model
        except ValueError: raise
        except Exception as e: errors.append(f"[{m}] {e}")
    raise ValueError("Groq failed:\n" + "\n".join(errors))


def stream_chat(api_key: str, system: str, user: str,
                max_tokens: int = 8000, model: str = "llama-3.3-70b-versatile"):
    """Streaming Groq call — yields text chunks."""
    models = [model] + [m for m in ALL_MODELS if m != model]
    for m in models:
        try:
            r = requests.post(
                GROQ_BASE,
                headers={"Authorization": f"Bearer {api_key}",
                         "Content-Type": "application/json"},
                json={**_body(system, user, max_tokens, m), "stream": True},
                stream=True, timeout=120
            )
            if r.status_code == 401:
                raise ValueError("Invalid Groq API key.")
            if r.status_code != 200:
                continue

            for raw in r.iter_lines(decode_unicode=True):
                if not raw or not raw.startswith("data: "):
                    continue
                chunk = raw[6:]
                if chunk.strip() == "[DONE]":
                    return
                try:
                    delta = json.loads(chunk)["choices"][0]["delta"].get("content","")
                    if delta:
                        yield delta
                except Exception:
                    pass
            return  # success

        except ValueError: raise
        except Exception: continue

    # Fallback to blocking
    yield chat(api_key, system, user, max_tokens, model)


def _repair_truncated_json(text: str) -> str:
    """Close unclosed brackets caused by token limits."""
    stack, in_string, escape_next = [], False, False
    for ch in text:
        if escape_next:       escape_next = False; continue
        if ch == '\\' and in_string: escape_next = True; continue
        if ch == '"':         in_string = not in_string; continue
        if in_string:         continue
        if ch in ('{', '['):  stack.append(ch)
        elif ch == '}' and stack and stack[-1] == '{': stack.pop()
        elif ch == ']' and stack and stack[-1] == '[': stack.pop()
    repaired = text.rstrip().rstrip(',')
    for ch in reversed(stack):
        repaired += '}' if ch == '{' else ']'
    return repaired


def parse_json_from_response(text: str) -> list:
    """Robust JSON extractor: handles fences, preamble, truncation, trailing commas."""
    if not text:
        raise ValueError("Empty response")

    cleaned = re.sub(r"```+\w*", "", text).strip()

    def _try_parse(s):
        s = re.sub(r",\s*([\]\}])", r"\1", s)   # fix trailing commas
        try:
            r = json.loads(s)
            return r if isinstance(r, list) else [r]
        except Exception:
            pass
        repaired = _repair_truncated_json(s)
        repaired = re.sub(r",\s*([\]\}])", r"\1", repaired)
        try:
            r = json.loads(repaired)
            return r if isinstance(r, list) else [r]
        except Exception:
            return None

    # Try [ ... ]
    s, e = cleaned.find("["), cleaned.rfind("]")
    if s != -1:
        result = _try_parse(cleaned[s: e+1] if e > s else cleaned[s:])
        if result is not None:
            return result

    # Try { ... }
    s, e = cleaned.find("{"), cleaned.rfind("}")
    if s != -1:
        result = _try_parse(cleaned[s: e+1] if e > s else cleaned[s:])
        if result is not None:
            return result

    raise ValueError(f"Could not extract valid JSON.\nFirst 500 chars:\n{text[:500]}")

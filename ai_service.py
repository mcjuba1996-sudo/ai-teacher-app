"""Small Gemini adapter with per-request credentials and safe UI errors.

No secret or provider error body is included in an AIError. The UI should
translate its ``code`` rather than display an underlying SDK exception.
"""

from collections.abc import Mapping, Sequence
import json
import time
from typing import Any

from google import genai
from google.genai import errors, types


DEFAULT_MODEL = "gemini-3.6-flash"
REQUEST_TIMEOUT_MS = 60_000
RETRY_DELAY_SECONDS = 1.0
_ERROR_CODES = frozenset(
    {"missing_key", "quota", "auth", "unavailable", "timeout", "blocked",
     "empty", "invalid_response", "request"}
)
_BLOCKED_FINISH_REASONS = frozenset(
    {"SAFETY", "RECITATION", "BLOCKLIST", "PROHIBITED_CONTENT", "SPII",
     "IMAGE_SAFETY", "IMAGE_PROHIBITED_CONTENT"}
)


class AIError(Exception):
    """An error code that is safe to expose to an end user."""

    def __init__(self, code: str):
        self.code = code if code in _ERROR_CODES else "request"
        super().__init__(self.code)


def _setting(settings: Mapping[str, Any], name: str) -> str:
    value = settings.get(name, "")
    return value.strip() if isinstance(value, str) else ""


def resolve_settings(
    env: Mapping[str, Any],
    secrets: Mapping[str, Any],
    user_key: str = "",
) -> tuple[str, str, bool]:
    """Return (key, model, using_shared) without reading global settings.

    A personal key wins over environment variables, then Streamlit secrets.
    DEFAULT_API_KEY remains supported for existing Streamlit deployments.
    """
    personal_key = user_key.strip() if isinstance(user_key, str) else ""
    shared_key = next(
        (value for settings, names in (
            (env, ("GEMINI_API_KEY", "GOOGLE_API_KEY", "DEFAULT_API_KEY")),
            (secrets, ("DEFAULT_API_KEY", "GEMINI_API_KEY", "GOOGLE_API_KEY")),
        ) for name in names if (value := _setting(settings, name))),
        "",
    )
    model = _setting(env, "GEMINI_MODEL") or _setting(secrets, "GEMINI_MODEL") or DEFAULT_MODEL
    return personal_key or shared_key, model, bool(shared_key and not personal_key)


def _enum_name(value: Any) -> str:
    """Read SDK enums and simple string test doubles without provider messages."""
    return str(getattr(value, "value", value) or "").rsplit(".", 1)[-1].upper()


def _api_error_code(exc: errors.APIError) -> str:
    status = _enum_name(getattr(exc, "status", None))
    code = getattr(exc, "code", None)
    if code in (401, 403) or status in ("UNAUTHENTICATED", "PERMISSION_DENIED"):
        return "auth"
    # Invalid API keys can use HTTP 400, with a structured ErrorInfo reason.
    body = getattr(exc, "details", None)
    if isinstance(body, Mapping):
        error = body.get("error", body)
        if isinstance(error, Mapping):
            details = error.get("details", [])
            if isinstance(details, list) and any(
                isinstance(item, Mapping) and item.get("reason") in
                {"API_KEY_INVALID", "API_KEY_EXPIRED", "API_KEY_SERVICE_BLOCKED"}
                for item in details
            ):
                return "auth"
    if code == 429 or status == "RESOURCE_EXHAUSTED":
        return "quota"
    if code in (408, 504) or status == "DEADLINE_EXCEEDED":
        return "timeout"
    if code in (404, 500, 502, 503) or status == "UNAVAILABLE":
        return "unavailable"
    return "request"


def _transport_error_code(exc: Exception) -> str | None:
    if isinstance(exc, TimeoutError):
        return "timeout"
    if isinstance(exc, ConnectionError):
        return "unavailable"
    # SDK releases support httpx and httpx2. Avoid requiring both transports.
    bases = type(exc).__mro__
    if any(base.__module__.split(".")[0] in {"httpx", "httpx2"} for base in bases):
        names = {base.__name__ for base in bases}
        if "TimeoutException" in names:
            return "timeout"
        if "NetworkError" in names or "RemoteProtocolError" in names:
            return "unavailable"
    return None


def _response_text(response: Any, json_mode: bool) -> str:
    feedback = getattr(response, "prompt_feedback", None)
    reason = _enum_name(getattr(feedback, "block_reason", None))
    if reason not in ("", "0", "BLOCK_REASON_UNSPECIFIED", "BLOCKED_REASON_UNSPECIFIED"):
        raise AIError("blocked")
    candidates = getattr(response, "candidates", None) or []
    for candidate in candidates:
        finish_reason = _enum_name(getattr(candidate, "finish_reason", None))
        if finish_reason in _BLOCKED_FINISH_REASONS:
            raise AIError("blocked")
        if finish_reason == "MAX_TOKENS":
            raise AIError("invalid_response")
    try:
        text = response.text
    except Exception:
        raise AIError("invalid_response") from None
    if not isinstance(text, str) or not text.strip():
        raise AIError("empty")
    text = text.strip()
    if json_mode:
        try:
            json.loads(text)
        except (TypeError, ValueError):
            raise AIError("invalid_response") from None
    return text


def generate_text(
    prompt: str,
    api_key: str,
    model: str = DEFAULT_MODEL,
    attachments: Sequence[tuple[str, bytes]] | None = None,
    json_mode: bool = False,
) -> str:
    """Generate one response; allow at most one retry for transient failures.

    ``attachments`` contains (MIME type, bytes) pairs. A new client is created
    and closed on every call, so concurrent Streamlit users never share keys.
    SDK retries are disabled to avoid nested retries and raw exception logs.
    Timeouts, quota limits, invalid credentials and invalid output are not retried.
    """
    if not isinstance(api_key, str) or not api_key.strip():
        raise AIError("missing_key")
    if not isinstance(prompt, str) or not prompt.strip():
        raise AIError("request")
    if not isinstance(model, str) or not model.strip():
        raise AIError("request")

    client = None
    try:
        parts = [types.Part.from_text(text=prompt)]
        for mime_type, data in attachments or []:
            if not isinstance(mime_type, str) or not mime_type.strip() or not isinstance(data, bytes) or not data:
                raise AIError("request")
            parts.append(types.Part.from_bytes(data=data, mime_type=mime_type))
        config = types.GenerateContentConfig(
            response_mime_type="application/json" if json_mode else "text/plain",
        )
        client = genai.Client(
            api_key=api_key.strip(),
            http_options=types.HttpOptions(
                timeout=REQUEST_TIMEOUT_MS,
                retry_options=types.HttpRetryOptions(attempts=1),
            ),
        )
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model.strip(),
                    contents=[types.Content(role="user", parts=parts)],
                    config=config,
                )
                return _response_text(response, json_mode)
            except errors.APIError as exc:
                error_code = _api_error_code(exc)
                transient = getattr(exc, "code", None) in (500, 502, 503)
                if attempt == 0 and transient:
                    time.sleep(RETRY_DELAY_SECONDS)
                    continue
                raise AIError(error_code) from None
            except AIError:
                raise
            except errors.UnknownApiResponseError:
                raise AIError("invalid_response") from None
            except Exception as exc:
                error_code = _transport_error_code(exc)
                if attempt == 0 and error_code == "unavailable":
                    time.sleep(RETRY_DELAY_SECONDS)
                    continue
                raise AIError(error_code or "request") from None
    except AIError:
        raise
    except errors.APIError as exc:
        raise AIError(_api_error_code(exc)) from None
    except Exception as exc:
        raise AIError(_transport_error_code(exc) or "request") from None
    finally:
        if client is not None:
            try:
                client.close()
            except Exception:
                # Cleanup errors must neither mask a response nor expose a key.
                pass
    raise AIError("unavailable")

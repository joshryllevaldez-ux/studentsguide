from __future__ import annotations

import asyncio
import copy
import hashlib
import html
import io
import json
import logging
import os
import re
import threading
import time
from collections import OrderedDict, defaultdict, deque
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from typing import Any

import httpx
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image as RLImage,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))
STATIC_DIR = os.path.join(BASE_DIR, "static")
ASSET_DIR = os.path.join(BASE_DIR, "assets")
MAX_UPLOAD_MB = max(1, int(os.getenv("MAX_UPLOAD_MB", "50")))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024
MAX_SOURCE_CHARS = 55_000
MAX_REQUEST_BYTES = (MAX_UPLOAD_MB + 2) * 1024 * 1024
IP_AI_GENERATION_LIMIT_PER_HOUR = int(os.getenv("IP_AI_GENERATION_LIMIT_PER_HOUR", "25"))
DAILY_AI_GENERATION_LIMIT = int(os.getenv("DAILY_AI_GENERATION_LIMIT", "100"))
AI_CONCURRENCY_LIMIT = max(1, int(os.getenv("AI_CONCURRENCY_LIMIT", "3")))
AI_CALL_TIMEOUT_SECONDS = max(45, int(os.getenv("AI_CALL_TIMEOUT_SECONDS", "120")))
AI_RESPONSE_CACHE_TTL_SECONDS = max(0, int(os.getenv("AI_RESPONSE_CACHE_TTL_SECONDS", "1800")))
AI_RESPONSE_CACHE_MAX_ENTRIES = max(0, int(os.getenv("AI_RESPONSE_CACHE_MAX_ENTRIES", "12")))
_AI_RESPONSE_CACHE: OrderedDict[str, tuple[float, dict[str, Any]]] = OrderedDict()
_AI_RESPONSE_CACHE_LOCK = threading.Lock()
_AI_CACHE_HITS = 0
_AI_CACHE_MISSES = 0
_AI_SEMAPHORE = asyncio.Semaphore(AI_CONCURRENCY_LIMIT)
_AI_LIMIT_LOCK = threading.Lock()
_AI_REQUESTS_BY_IP: dict[str, deque[float]] = defaultdict(deque)
_AI_DAILY_DATE = datetime.now(timezone.utc).date()
_AI_DAILY_COUNT = 0
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
if GEMINI_API_KEY.lower() in {"your_api_key_here", "replace_me", "changeme"}:
    GEMINI_API_KEY = ""
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
GEMINI_FAST_MODEL = os.getenv("GEMINI_FAST_MODEL", "gemini-3.5-flash-lite").strip()
GEMINI_BASE_URL = "https://generativelanguage.googleapis.com/v1beta"

APP_VERSION = "0.4.9"
PROMPT_VERSION = "1.2"
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
)
LOGGER = logging.getLogger("studentsguide")


@asynccontextmanager
async def lifespan(application: FastAPI):
    application.state.http_client = httpx.AsyncClient(
        timeout=httpx.Timeout(180.0, connect=20.0),
        limits=httpx.Limits(max_connections=8, max_keepalive_connections=4, keepalive_expiry=30.0),
    )
    LOGGER.info("service_started ai_model=%s timeout_seconds=%s max_upload_mb=%s concurrency=%s", GEMINI_FAST_MODEL, AI_CALL_TIMEOUT_SECONDS, MAX_UPLOAD_MB, AI_CONCURRENCY_LIMIT)
    try:
        yield
    finally:
        await application.state.http_client.aclose()


app = FastAPI(title="StudentsGUIDE", version=APP_VERSION, lifespan=lifespan)


@app.exception_handler(Exception)
async def unexpected_error_handler(request: Request, exc: Exception):
    LOGGER.exception("unhandled_api_error method=%s path=%s", request.method, request.url.path)
    return JSONResponse(
        status_code=500,
        content={"detail": "StudentsGUIDE encountered an unexpected server error. Check the Render service logs for details, then try again."},
    )


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
app.mount("/assets", StaticFiles(directory=ASSET_DIR), name="assets")


@app.middleware("http")
async def security_headers_and_request_limits(request: Request, call_next):
    content_length = request.headers.get("content-length")
    if content_length:
        try:
            if int(content_length) > MAX_REQUEST_BYTES:
                return PlainTextResponse("Request is too large.", status_code=413)
        except ValueError:
            return PlainTextResponse("Invalid request size header.", status_code=400)

    started = time.perf_counter()
    response = None
    try:
        response = await call_next(request)
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("X-Frame-Options", "DENY")
        response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
        response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
        return response
    finally:
        elapsed_ms = (time.perf_counter() - started) * 1000
        LOGGER.info(
            "http_request method=%s path=%s status=%s duration_ms=%.1f",
            request.method,
            request.url.path,
            response.status_code if response is not None else 500,
            elapsed_ms,
        )


def enforce_ai_generation_quota(request: Request) -> None:
    """Small single-instance abuse guard for a shared classroom demo deployment."""
    global _AI_DAILY_DATE, _AI_DAILY_COUNT
    now = time.time()
    today = datetime.now(timezone.utc).date()
    # Uvicorn's proxy-header support supplies the original client address on Render.
    client_ip = request.client.host if request.client else "unknown"
    with _AI_LIMIT_LOCK:
        if today != _AI_DAILY_DATE:
            _AI_DAILY_DATE = today
            _AI_DAILY_COUNT = 0
            _AI_REQUESTS_BY_IP.clear()
        requests = _AI_REQUESTS_BY_IP[client_ip]
        while requests and now - requests[0] >= 3600:
            requests.popleft()
        if len(requests) >= IP_AI_GENERATION_LIMIT_PER_HOUR:
            raise HTTPException(
                status_code=429,
                detail="This network has reached its hourly StudentsGUIDE AI limit. Please try again later.",
                headers={"Retry-After": "3600"},
            )
        if _AI_DAILY_COUNT >= DAILY_AI_GENERATION_LIMIT:
            raise HTTPException(
                status_code=429,
                detail="The shared StudentsGUIDE daily AI allowance has been reached. Please try again tomorrow.",
                headers={"Retry-After": "3600"},
            )
        requests.append(now)
        _AI_DAILY_COUNT += 1


class QuizRequest(BaseModel):
    source_text: str = Field(min_length=80, max_length=MAX_SOURCE_CHARS)
    mode: str
    lecture_title: str = "Uploaded lecture"


class PdfRequest(BaseModel):
    guide: dict[str, Any]


def _iter_slide_shapes(shapes):
    """Yield top-level and grouped PowerPoint shapes in their slide order."""
    for shape in shapes:
        yield shape
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_slide_shapes(shape.shapes)


def _extract_shape_text(shape) -> list[str]:
    """Extract readable text from a shape without attempting OCR."""
    values: list[str] = []
    if getattr(shape, "has_text_frame", False):
        text = "\n".join(paragraph.text for paragraph in shape.text_frame.paragraphs).strip()
        if text:
            values.append(text)
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
                values.append(row_text)
    if getattr(shape, "has_chart", False):
        try:
            chart = shape.chart
            categories: list[str] = []
            try:
                categories = [str(category.label) for category in chart.plots[0].categories]
            except Exception:
                categories = []
            for series in chart.series:
                series_name = str(getattr(series, "name", "Series")).strip() or "Series"
                for index, value in enumerate(series.values):
                    category = categories[index] if index < len(categories) else f"Item {index + 1}"
                    values.append(f"Chart · {series_name} · {category}: {value}")
        except Exception:
            # A malformed or unsupported chart must not prevent other slides being read.
            pass
    return values


def extract_pptx_text(data: bytes) -> tuple[str, int]:
    try:
        prs = Presentation(io.BytesIO(data))
    except Exception as exc:
        raise HTTPException(status_code=400, detail="That file could not be opened as a PowerPoint presentation. Please upload a valid .pptx file.") from exc

    slide_blocks: list[str] = []
    for index, slide in enumerate(prs.slides, start=1):
        lines: list[str] = []
        for shape in _iter_slide_shapes(slide.shapes):
            lines.extend(_extract_shape_text(shape))

        try:
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes and notes.lower() not in {"slide number", str(index)}:
                    lines.append("Speaker notes: " + notes)
        except Exception:
            pass

        # Remove exact duplicates while retaining the original reading order.
        unique_lines = list(dict.fromkeys(line.strip() for line in lines if line.strip()))
        if unique_lines:
            slide_blocks.append(f"SLIDE {index}\n" + "\n".join(unique_lines))
        else:
            slide_blocks.append(
                f"SLIDE {index}\n[NO SELECTABLE TEXT FOUND: this slide may contain image-only content. "
                "Visual text inside images has not been interpreted.]"
            )

    text = "\n\n".join(slide_blocks).strip()
    if len(text) < 80:
        raise HTTPException(status_code=422, detail="I could not find enough readable text in this PowerPoint. It may consist mainly of screenshots or image-only slides, which are not interpreted automatically.")
    if len(text) > MAX_SOURCE_CHARS:
        raise HTTPException(
            status_code=413,
            detail=(
                f"This PowerPoint contains {len(text):,} characters of extracted lecture text, "
                f"which exceeds the current safe processing limit of {MAX_SOURCE_CHARS:,}. "
                "No study guide was generated because doing so would omit lecture content. "
                "Please split the presentation into smaller parts and upload them separately."
            ),
        )
    return text, len(prs.slides)

def extract_output_text(payload: dict[str, Any]) -> str:
    direct = payload.get("output_text")
    if isinstance(direct, str) and direct.strip():
        return direct.strip()
    chunks: list[str] = []
    for item in payload.get("output", []) or []:
        if item.get("type") == "message":
            for content in item.get("content", []) or []:
                if content.get("type") in ("output_text", "text") and content.get("text"):
                    chunks.append(content["text"])
    return "\n".join(chunks).strip()


def parse_json_text(text: str) -> dict[str, Any]:
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.I)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        obj = json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, flags=re.S)
        if not match:
            raise HTTPException(status_code=502, detail="The AI response was not in the expected structured format. Please try generating again.")
        try:
            obj = json.loads(match.group(0))
        except json.JSONDecodeError as exc:
            raise HTTPException(status_code=502, detail="The AI response could not be read. Please try generating again.") from exc
    if not isinstance(obj, dict):
        raise HTTPException(status_code=502, detail="The AI response did not contain a valid result object.")
    return obj


async def _ask_ai_uncached(
    instructions: str,
    user_input: str,
    max_output_tokens: int = 6500,
    operation: str = "generation",
) -> dict[str, Any]:
    if not GEMINI_API_KEY:
        raise HTTPException(
            status_code=503,
            detail="Gemini AI is not connected. The administrator must configure GEMINI_API_KEY on the hosting server.",
        )

    fast_quiz_operations = {"quiz_mcq_20", "quiz_mcq_10", "quiz_true_or_false_10"}
    use_fast_model = operation == "study_guide" or operation in fast_quiz_operations
    model = GEMINI_FAST_MODEL  # Use the high-throughput model for every mode to avoid the 20-request/day model quota.
    thinking_level = "minimal" if use_fast_model else "low"
    output_limit = min(max(1024, int(max_output_tokens)), 65536)
    body = {
        "systemInstruction": {"parts": [{"text": instructions}]},
        "contents": [{
            "role": "user",
            "parts": [{
                "text": (
                    "Return exactly one valid JSON object. Do not use Markdown fences. "
                    "The response must be valid JSON.\n\n" + user_input
                )
            }]
        }],
        "generationConfig": {
            "responseMimeType": "application/json",
            # Low thinking reduces latency; maxOutputTokens still includes thinking tokens.
            "thinkingConfig": {"thinkingLevel": thinking_level},
            "maxOutputTokens": output_limit,
        },
    }
    headers = {"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"}
    started = time.perf_counter()

    async def send_with_bounded_retry(active_client: httpx.AsyncClient) -> httpx.Response:
        response = None
        for attempt in range(3):
            response = await active_client.post(
                f"{GEMINI_BASE_URL}/models/{model}:generateContent",
                headers=headers,
                json=body,
            )
            # Retry only transient upstream server failures. Do not retry quota,
            # authentication, validation, or client errors.
            if response.status_code not in {500, 502, 503, 504} or attempt == 2:
                return response
            delay = 0.8 * (2 ** attempt)
            LOGGER.warning(
                "gemini_transient_status operation=%s status=%s attempt=%s retry_in_seconds=%.1f",
                operation, response.status_code, attempt + 1, delay,
            )
            await asyncio.sleep(delay)
        return response

    try:
        async with asyncio.timeout(AI_CALL_TIMEOUT_SECONDS):
            async with _AI_SEMAPHORE:
                client = getattr(app.state, "http_client", None)
                if client is None:
                    # Allows local scripts/tests that invoke the function without ASGI lifespan.
                    async with httpx.AsyncClient(timeout=httpx.Timeout(125.0, connect=15.0)) as fallback:
                        response = await send_with_bounded_retry(fallback)
                        if response.status_code < 400 and output_limit < 65536:
                            initial_payload = response.json()
                            initial_candidates = initial_payload.get("candidates", [])
                            if initial_candidates and initial_candidates[0].get("finishReason") == "MAX_TOKENS":
                                LOGGER.warning(
                                    "gemini_length_recovery operation=%s model=%s initial_budget=%s retry_budget=65536",
                                    operation, model, output_limit,
                                )
                                output_limit = 65536
                                body["generationConfig"]["maxOutputTokens"] = output_limit
                                response = await send_with_bounded_retry(fallback)
                else:
                    response = await send_with_bounded_retry(client)
                    if response.status_code < 400 and output_limit < 65536:
                        initial_payload = response.json()
                        initial_candidates = initial_payload.get("candidates", [])
                        if initial_candidates and initial_candidates[0].get("finishReason") == "MAX_TOKENS":
                            LOGGER.warning(
                                "gemini_length_recovery operation=%s model=%s initial_budget=%s retry_budget=65536",
                                operation, model, output_limit,
                            )
                            output_limit = 65536
                            body["generationConfig"]["maxOutputTokens"] = output_limit
                            response = await send_with_bounded_retry(client)
    except TimeoutError as exc:
        LOGGER.warning("gemini_total_timeout operation=%s model=%s limit_seconds=%s duration_ms=%.1f", operation, model, AI_CALL_TIMEOUT_SECONDS, (time.perf_counter() - started) * 1000)
        raise HTTPException(status_code=504, detail=f"AI generation exceeded the {AI_CALL_TIMEOUT_SECONDS}-second safety limit. Please retry; if this repeats, check Render logs for the model response time.") from exc
    except httpx.TimeoutException as exc:
        LOGGER.warning("gemini_network_timeout operation=%s duration_ms=%.1f", operation, (time.perf_counter() - started) * 1000)
        raise HTTPException(status_code=504, detail="Gemini network request timed out. Please retry.") from exc
    except httpx.HTTPError as exc:
        LOGGER.exception("gemini_connection_error operation=%s", operation)
        raise HTTPException(status_code=502, detail="StudentsGUIDE could not reach Gemini. Check the server connection.") from exc

    elapsed_ms = (time.perf_counter() - started) * 1000
    if response.status_code >= 400:
        try:
            error = response.json().get("error", {})
            detail = error.get("message", "Gemini rejected the request.")
        except Exception:
            detail = "Gemini rejected the request."
        LOGGER.warning(
            "gemini_error operation=%s model=%s status=%s duration_ms=%.1f input_chars=%s detail=%s",
            operation, model, response.status_code, elapsed_ms, len(user_input), detail[:300],
        )
        status_code = 429 if response.status_code == 429 else (503 if response.status_code == 503 else 502)
        raise HTTPException(status_code=status_code, detail=f"Gemini service error: {detail}")

    payload = response.json()
    candidates = payload.get("candidates", [])
    if not candidates:
        LOGGER.warning("gemini_empty_candidates operation=%s duration_ms=%.1f", operation, elapsed_ms)
        raise HTTPException(status_code=502, detail="Gemini returned no response. Please try again.")

    candidate = candidates[0]
    finish_reason = candidate.get("finishReason", "UNKNOWN")
    parts = candidate.get("content", {}).get("parts", [])
    output = "\n".join(
        part.get("text", "")
        for part in parts
        if isinstance(part, dict) and part.get("text") and not part.get("thought", False)
    ).strip()
    usage = payload.get("usageMetadata", {})
    LOGGER.info(
        "gemini_complete operation=%s model=%s duration_ms=%.1f input_chars=%s output_chars=%s "
        "finish_reason=%s output_budget=%s prompt_tokens=%s output_tokens=%s thinking_tokens=%s",
        operation, model, elapsed_ms, len(user_input), len(output), finish_reason, output_limit,
        usage.get("promptTokenCount"), usage.get("candidatesTokenCount"), usage.get("thoughtsTokenCount"),
    )

    if finish_reason == "MAX_TOKENS":
        raise HTTPException(
            status_code=502,
            detail=(
                "Gemini still reached its 65,536-token output ceiling after an automatic full-budget retry. "
                "No incomplete study guide was returned. Split this PowerPoint into smaller lecture parts and retry; "
                "the Render log records the model, output budget and token usage."
            ),
        )
    if not output:
        raise HTTPException(status_code=502, detail="Gemini returned an empty response. Please try again.")
    return parse_json_text(output)


def _ai_cache_key(instructions: str, user_input: str, max_output_tokens: int, operation: str) -> str:
    """Hash a complete AI request without storing source text in the cache key."""
    material = "\\0".join((
        GEMINI_FAST_MODEL, PROMPT_VERSION, operation, str(max_output_tokens),
        instructions, user_input,
    ))
    return hashlib.sha256(material.encode("utf-8")).hexdigest()


def _read_ai_response_cache(key: str) -> dict[str, Any] | None:
    """Return a defensive copy of a fresh in-memory response, if available."""
    global _AI_CACHE_HITS, _AI_CACHE_MISSES
    if AI_RESPONSE_CACHE_TTL_SECONDS <= 0 or AI_RESPONSE_CACHE_MAX_ENTRIES <= 0:
        return None
    now = time.monotonic()
    with _AI_RESPONSE_CACHE_LOCK:
        entry = _AI_RESPONSE_CACHE.get(key)
        if entry is None or now - entry[0] >= AI_RESPONSE_CACHE_TTL_SECONDS:
            if entry is not None:
                _AI_RESPONSE_CACHE.pop(key, None)
            _AI_CACHE_MISSES += 1
            return None
        _AI_RESPONSE_CACHE.move_to_end(key)
        _AI_CACHE_HITS += 1
        return copy.deepcopy(entry[1])


def _store_ai_response_cache(key: str, value: dict[str, Any]) -> None:
    if AI_RESPONSE_CACHE_TTL_SECONDS <= 0 or AI_RESPONSE_CACHE_MAX_ENTRIES <= 0:
        return
    now = time.monotonic()
    with _AI_RESPONSE_CACHE_LOCK:
        _AI_RESPONSE_CACHE[key] = (now, copy.deepcopy(value))
        _AI_RESPONSE_CACHE.move_to_end(key)
        while len(_AI_RESPONSE_CACHE) > AI_RESPONSE_CACHE_MAX_ENTRIES:
            _AI_RESPONSE_CACHE.popitem(last=False)
        expired = [
            cache_key for cache_key, (created, _) in _AI_RESPONSE_CACHE.items()
            if now - created >= AI_RESPONSE_CACHE_TTL_SECONDS
        ]
        for cache_key in expired:
            _AI_RESPONSE_CACHE.pop(cache_key, None)


async def ask_ai(
    instructions: str,
    user_input: str,
    max_output_tokens: int = 6500,
    operation: str = "generation",
    request: Request | None = None,
) -> dict[str, Any]:
    """Reuse recent identical generations before consuming quota or calling Gemini."""
    cache_key = _ai_cache_key(instructions, user_input, max_output_tokens, operation)
    cached = _read_ai_response_cache(cache_key)
    if cached is not None:
        LOGGER.info("ai_cache_hit operation=%s cache_key=%s", operation, cache_key[:12])
        return cached

    # The app's abuse guard only applies when a real provider request is required.
    if request is not None:
        enforce_ai_generation_quota(request)

    result = await _ask_ai_uncached(
        instructions,
        user_input,
        max_output_tokens=max_output_tokens,
        operation=operation,
    )
    _store_ai_response_cache(cache_key, result)
    return result


STUDY_GUIDE_INSTRUCTIONS = """You are StudentsGUIDE, an expert university-level tutor and meticulous lecture-to-learning designer. Your job is to transform the supplied PowerPoint text into a genuinely useful, comprehensive study guide that helps a student understand, remember, and apply the lecture—not merely summarize it.

SOURCE FIDELITY AND COVERAGE
- Use the supplied lecture as the authority. Do not invent lecture facts, definitions, numbers, formulas, case-study outcomes, citations, or claims.
- You may add a clearly labeled simple analogy or everyday example to clarify a source concept, but it must not introduce a new factual claim or change the concept.
- Preserve the lecturer's terminology, named models, examples, case studies, distinctions, processes, and important qualifications.
- Work through the slide-organized source from beginning to end. Map every substantive slide/topic to one or more sections. Combine closely related slides when useful, but do not silently omit important content to keep the guide short.
- Include source slide numbers in slide_refs. If a slide contains only a diagram/image whose contents are not present in the supplied text, do not pretend to understand it; mention this limitation in coverage_note.
- If the source is ambiguous, contradictory, incomplete, or appears to contain a questionable claim, flag it neutrally in coverage_note instead of silently correcting it.
- Do not claim full coverage unless the content actually supports it.

TEACHING STANDARD
- Organize the guide in lecture order. Use one section per coherent topic, normally 4-10 sections. Combine only closely related topics; do not omit examinable content.
- Each section should teach a coherent topic. Use clear headings and preserve the lecture's logical order.
- Explain concepts in plain, student-friendly language first, then retain the correct technical terminology.
- For each important concept, explain what it is and how it works. Include why it matters only when the lecture supports it; do not repeat the same point in multiple fields.
- Break processes, workflows, formulas, and mechanisms into numbered steps. Explain what happens at each step and why.
- Explain comparisons explicitly: show what is different, what is similar, and why the distinction matters. Do not invent comparison facts absent from the lecture.
- Include one concise worked example when the source supports one. Add an extra example only for a distinct calculation or concept that would otherwise be unclear.
- For calculations or formulas, preserve exact source notation and values, define variables when provided, show each calculation step, and explain the result. If the source contains no calculation, do not invent one.
- Include common mistakes or misconceptions only when they can be derived safely from the source; label them as study cautions, not as claims made by the lecturer.
- Brainhacks must be genuinely useful retrieval cues: mnemonic, keyword association, analogy, sequence cue, or memory formula. Explain what each part of the cue means. Never use a catchy cue that distorts the facts.
- Keep fields compact and non-repetitive: 2-4 key points, 1-3 process steps, up to 4 important terms, and up to 3 exam-focus items per section. Include one self-check question and answer for each section.
- Highlight exam focus: definitions, distinctions, steps, named models, case-study facts, formulas, and likely application points supported by the lecture.
- Avoid filler, repetitive introductions, vague praise, and overly compressed bullet lists. Prefer meaningful explanation over artificial brevity.

Return ONLY valid JSON, no Markdown fences, with this shape:
{
  "title": "short lecture title",
  "course": "course or subject if inferable, otherwise Not specified",
  "overview": "clear 3-5 sentence overview of what the lecture teaches and how its topics connect",
  "learning_objectives": ["specific measurable things a student should be able to explain, distinguish, or apply"],
  "sections": [
    {
      "heading": "specific topic heading",
      "slide_refs": [1, 2],
      "key_points": ["important source-grounded point"],
      "explanation": "substantial plain-language teaching explanation, not a one-line summary",
      "why_it_matters": "why this concept matters in the lecture or its application, grounded in the source",
      "how_it_works": ["numbered logical step or mechanism; use an empty array if not applicable"],
      "example": {
        "title": "Worked example",
        "scenario": "specific, understandable situation grounded in the lecture",
        "steps": ["step-by-step reasoning and explanation"],
        "takeaway": "what the student should learn from this example"
      },
      "additional_examples": [
        {"title": "Second example", "scenario": "another distinct example if useful and supported", "steps": ["reasoning"], "takeaway": "lesson"}
      ],
      "comparison": "explain a relevant comparison from the source, or empty string",
      "common_mistakes": ["a likely confusion and how to avoid it, based on source distinctions"],
      "brainhack": "a memorable cue followed by a brief explanation of what it means",
      "important_terms": [{"term": "technical term", "meaning": "accurate plain-language definition"}],
      "self_check": {"question": "active-recall question", "answer": "clear answer"},
      "exam_focus": ["specific detail, distinction, process, or application worth revising"]
    }
  ],
  "quick_recap": ["high-value takeaways across the entire lecture"],
  "exam_tips": ["specific revision advice grounded in the lecture"],
  "coverage_note": "state whether substantive slide topics were covered; identify any missing text, image-only slides, unclear details, or truncation honestly"
}
Return a complete but efficient guide. Cover every substantive text-based slide exactly once or in a clearly related section. Keep each explanation around 50-80 words, why_it_matters to one sentence, worked examples to 2-4 concise steps, and optional lists short. Use only facts and examples supported by the uploaded lecture. If a diagram or image has no extractable text, disclose that limitation. Never invent missing content."""


QUIZ_INSTRUCTIONS = """You are StudentsGUIDE's assessment generator. Generate questions ONLY from the supplied lecture. Do not invent facts or test outside knowledge. Keep exact requested counts and types. Avoid repeated concepts.

Return ONLY one valid JSON object:
{"quiz_title":"...","mode":"...","instructions":"...","questions":[{"id":1,"type":"mcq|true_false|scenario","topic":"...","context":"short scenario context or empty string","question":"...","options":[{"label":"A","text":"..."},{"label":"B","text":"..."},{"label":"C","text":"..."},{"label":"D","text":"..."}],"correct_answer":"A|B|C|D|True|False","explanation":"clear reason the answer is correct","example":"short illustrative example","brainhack":"brief accurate memory cue","source_basis":"lecture idea supporting the answer"}]}

True/false questions must have exactly True and False options and use correct_answer True or False. MCQ and scenario questions must have four plausible options and use the correct option label A-D. Scenarios must test application, not merely recall.

Keep each question concise but useful: question generally under 35 words; explanation 25-50 words; example one sentence (the first question may use two); brainhack under 18 words; source_basis under 15 words. Do not omit explanations or memory cues. Return exactly the requested number of questions. For MIXED 40, return exactly 16 MCQ, 12 true_false, and 12 scenario questions. Do not add introductory essays or extra fields."""


@app.get("/")
async def home() -> FileResponse:
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


@app.get("/api/status")
async def status() -> dict[str, Any]:
    return {
        "ok": True,
        "app_version": APP_VERSION,
        "ai_configured": bool(GEMINI_API_KEY),
        "model": GEMINI_FAST_MODEL if GEMINI_API_KEY else None,
        "study_model": GEMINI_FAST_MODEL if GEMINI_API_KEY else None,
        "applied_quiz_model": GEMINI_FAST_MODEL if GEMINI_API_KEY else None,
        "generation_timeout_seconds": AI_CALL_TIMEOUT_SECONDS,
        "max_upload_mb": MAX_UPLOAD_MB,
        "supported_file": ".pptx",
        "max_ai_generations_per_day": DAILY_AI_GENERATION_LIMIT,
        "ip_ai_generations_per_hour": IP_AI_GENERATION_LIMIT_PER_HOUR,
        "ai_cache_enabled": AI_RESPONSE_CACHE_TTL_SECONDS > 0 and AI_RESPONSE_CACHE_MAX_ENTRIES > 0,
        "ai_cache_ttl_seconds": AI_RESPONSE_CACHE_TTL_SECONDS,
        "ai_cache_max_entries": AI_RESPONSE_CACHE_MAX_ENTRIES,
        "ai_cache_entries": len(_AI_RESPONSE_CACHE),
        "ai_cache_hits": _AI_CACHE_HITS,
        "ai_cache_misses": _AI_CACHE_MISSES,
    }


@app.post("/api/generate-study-guide")
async def generate_study_guide(request: Request, file: UploadFile = File(...)) -> dict[str, Any]:
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=503, detail="AI generation is not connected yet. The StudentsGUIDE administrator must configure the API key on the hosting server.")
    filename = file.filename or "lecture.pptx"
    if not filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Please upload a PowerPoint .pptx file.")
    data = await file.read()
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail=f"The file is larger than {MAX_UPLOAD_MB} MB. Please upload a smaller PowerPoint.")
    if not data:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")

    extraction_started = time.perf_counter()
    source_text, slide_count = await asyncio.to_thread(extract_pptx_text, data)
    extraction_ms = (time.perf_counter() - extraction_started) * 1000
    LOGGER.info(
        "pptx_extracted filename=%s bytes=%s slides=%s source_chars=%s duration_ms=%.1f",
        filename, len(data), slide_count, len(source_text), extraction_ms,
    )
    prompt = (
        f"Create a source-faithful study guide. File: {filename}. Slide count: {slide_count}.\n"
        "Preserve all examinable concepts, definitions, formulas, classifications and lecture examples. "
        "Use concise teaching explanations and avoid repeating the same point.\n\n"
        "LECTURE CONTENT BY SLIDE:\n" + source_text
    )
    guide_output_budget = min(32000, 18000 + (len(source_text) // 5000) * 1200)
    guide = await ask_ai(
        STUDY_GUIDE_INSTRUCTIONS,
        prompt,
        max_output_tokens=guide_output_budget,
        operation="study_guide",
        request=request,
    )
    if not isinstance(guide.get("sections"), list) or not guide.get("sections"):
        raise HTTPException(status_code=502, detail="The AI response did not contain any study-guide topics. Please try generating again.")
    guide["title"] = str(guide.get("title") or os.path.splitext(filename)[0])
    guide["source_filename"] = filename
    guide["slide_count"] = slide_count
    guide["source_text"] = source_text
    guide["generation_type"] = "AI-generated from uploaded lecture"
    guide["prompt_version"] = PROMPT_VERSION
    return guide


MODE_SPECS = {
    "MCQ 20": (20, "Create 20 multiple-choice questions. All 20 must be mcq."),
    "MCQ 10": (10, "Create 10 multiple-choice questions. All 10 must be mcq."),
    "TRUE OR FALSE 10": (10, "Create 10 true_false questions. All 10 must be true_false."),
    "SCENARIO QUESTIONS 15": (15, "Create 15 scenario questions. All 15 must be scenario, each with an applied situation and four options."),
    "MIXED 40": (40, "Create 40 questions: 16 mcq, 12 true_false, and 12 scenario questions. Mix the question types throughout the quiz.")
}


@app.post("/api/generate-quiz")
async def generate_quiz(payload: QuizRequest, request: Request) -> dict[str, Any]:
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=503, detail="AI generation is not connected yet. The StudentsGUIDE administrator must configure the API key on the hosting server.")
    mode = payload.mode.upper().strip()
    if mode not in MODE_SPECS:
        raise HTTPException(status_code=400, detail="Choose one of the five available quiz modes.")
    count, mix_instruction = MODE_SPECS[mode]
    prompt = (
        f"Lecture title: {payload.lecture_title}\nSelected mode: {mode}\n"
        f"Requirement: {mix_instruction}\nExact number of questions: {count}.\n\n"
        "LECTURE CONTENT:\n" + payload.source_text[:MAX_SOURCE_CHARS]
    )
    quiz_budgets = {
        "MCQ 20": 9000,
        "MCQ 10": 5000,
        "TRUE OR FALSE 10": 3000,
        "SCENARIO QUESTIONS 15": 9000,
        "MIXED 40": 16000,
    }
    quiz = await ask_ai(
        QUIZ_INSTRUCTIONS,
        prompt,
        max_output_tokens=quiz_budgets[mode],
        operation=f"quiz_{mode.lower().replace(chr(32), chr(95))}",
        request=request,
    )
    questions = quiz.get("questions")
    if not isinstance(questions, list) or not questions:
        raise HTTPException(status_code=502, detail="The AI did not return a usable question list. Please generate the quiz again.")
    # Normalize IDs and remove malformed question entries rather than rendering broken controls.
    clean: list[dict[str, Any]] = []
    requested_type = {"MCQ 20": "mcq", "MCQ 10": "mcq", "TRUE OR FALSE 10": "true_false", "SCENARIO QUESTIONS 15": "scenario"}.get(mode)
    for item in questions:
        if not isinstance(item, dict) or not item.get("question") or not item.get("correct_answer"):
            continue
        item["type"] = str(item.get("type", "mcq")).lower()
        if requested_type and item["type"] != requested_type:
            continue
        if item["type"] not in ("mcq", "true_false", "scenario"):
            continue
        item["id"] = len(clean) + 1
        item.setdefault("explanation", "Review the related lecture topic and compare it with the selected answer.")
        item.setdefault("example", "Review the related lecture example.")
        item.setdefault("brainhack", "Connect the key term to its meaning in the lecture.")
        options = item.get("options") if isinstance(item.get("options"), list) else []
        if item["type"] == "true_false":
            item["options"] = [{"label": "True", "text": "True"}, {"label": "False", "text": "False"}]
            answer = str(item.get("correct_answer", "")).strip().lower()
            if answer in ("t", "true"):
                item["correct_answer"] = "True"
            elif answer in ("f", "false"):
                item["correct_answer"] = "False"
            else:
                continue
        else:
            normalized_options = []
            for idx, option in enumerate(options[:4]):
                if isinstance(option, dict) and option.get("text"):
                    normalized_options.append({"label": chr(65 + idx), "text": str(option["text"])})
            if len(normalized_options) != 4:
                continue
            item["options"] = normalized_options
            answer = str(item.get("correct_answer", "")).strip().upper()
            if answer not in ("A", "B", "C", "D"):
                continue
            item["correct_answer"] = answer
        clean.append(item)
    if mode == "MIXED 40":
        type_counts = {kind: sum(1 for q in clean if q["type"] == kind) for kind in ("mcq", "true_false", "scenario")}
        expected_counts = {"mcq": 16, "true_false": 12, "scenario": 12}
        if type_counts != expected_counts:
            raise HTTPException(
                status_code=502,
                detail=(
                    "The mixed quiz did not meet the required balance "
                    "(16 MCQ, 12 True/False and 12 Scenario questions). "
                    "Please generate it again."
                ),
            )
    if len(clean) < count:
        raise HTTPException(status_code=502, detail=f"The AI returned {len(clean)} usable questions instead of {count}. Please try generating again.")
    quiz["questions"] = clean[:count]
    quiz["mode"] = mode
    quiz["lecture_title"] = payload.lecture_title
    quiz["generation_type"] = "AI-generated from uploaded lecture"
    quiz["prompt_version"] = PROMPT_VERSION
    return quiz


def make_pdf(guide: dict[str, Any]) -> bytes:
    """Create a clean, topic-first study guide with a premium navy academic design."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=18*mm, leftMargin=18*mm,
        topMargin=17*mm, bottomMargin=19*mm,
        title=str(guide.get("title", "StudentsGUIDE Study Guide")),
        author="StudentsGUIDE.com",
    )
    navy = colors.HexColor("#102D69")
    navy_deep = colors.HexColor("#0B214F")
    blue = colors.HexColor("#246BD2")
    pale_blue = colors.HexColor("#EEF5FF")
    page_blue = colors.HexColor("#F3F7FD")
    gold = colors.HexColor("#FFF3C9")
    ink = colors.HexColor("#243653")
    muted = colors.HexColor("#657894")
    green = colors.HexColor("#EAF8F1")
    line = colors.HexColor("#D8E6F7")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SGTopic", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=16, leading=20, textColor=navy, spaceBefore=10, spaceAfter=8, keepWithNext=True))
    styles.add(ParagraphStyle(name="SGSubhead", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=9.5, leading=13, textColor=blue, spaceBefore=8, spaceAfter=4, keepWithNext=True))
    styles.add(ParagraphStyle(name="SGBodyNew", parent=styles["BodyText"], fontName="Helvetica", fontSize=9, leading=13.5, textColor=ink, spaceAfter=5))
    styles.add(ParagraphStyle(name="SGSmallNew", parent=styles["BodyText"], fontName="Helvetica", fontSize=8, leading=11.5, textColor=muted, spaceAfter=3))
    styles.add(ParagraphStyle(name="SGPanelWhite", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=12.2, textColor=colors.white, spaceAfter=2))
    styles.add(ParagraphStyle(name="SGPanelLabel", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=7.5, leading=9.5, textColor=colors.HexColor("#D7E8FF"), spaceAfter=3))
    styles.add(ParagraphStyle(name="SGGoldBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=12.2, textColor=colors.HexColor("#604A12"), spaceAfter=2))

    def P(value: Any, style: str = "SGBodyNew") -> Paragraph:
        # Decode escaped markup before stripping it, so legacy &lt;b&gt; tags
        # are removed too. All remaining text is escaped before ReportLab sees it.
        raw = html.unescape(str(value or ""))
        raw = re.sub(r"(?i)<br\s*/?>", "\n", raw)
        # Remove any remaining HTML/XML-style tags, not just bold tags.
        raw = re.sub(r"</?[A-Za-z][^>]*>", "", raw)
        escaped = html.escape(raw).replace("\n", "<br/>")
        return Paragraph(escaped, styles[style])

    def bullets(items: Any, style: str = "SGSmallNew") -> list[Any]:
        return [P("•  " + str(item), style) for item in (items or [])]

    def panel(label: str, body: Any, bg=navy, label_color=None, body_style="SGPanelWhite", accent=blue) -> Table:
        label_style = "SGPanelLabel" if label_color is None else label_color
        content = [[P(label.upper(), label_style)], [P(body, body_style)]]
        t = Table(content, colWidths=[doc.width], hAlign="LEFT")
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), bg),
            ("BOX", (0,0), (-1,-1), 0.55, accent),
            ("LINEBEFORE", (0,0), (0,-1), 4, accent),
            ("LEFTPADDING", (0,0), (-1,-1), 11),
            ("RIGHTPADDING", (0,0), (-1,-1), 11),
            ("TOPPADDING", (0,0), (-1,-1), 7),
            ("BOTTOMPADDING", (0,0), (-1,-1), 7),
        ]))
        return t

    story: list[Any] = []
    sections = guide.get("sections") or []
    # Begin immediately with the actual lecture topics. No logo, cover card,
    # promotional header, overview page, objectives panel, or roadmap.
    for index, section in enumerate(sections, start=1):
        heading = f"{index:02d}  {section.get('heading', 'Key topic')}"
        slide_refs = section.get("slide_refs") or []
        slide_label = "SLIDES " + ", ".join(map(str, slide_refs)) if slide_refs else "LECTURE TOPIC"
        head = Table([[P(heading, "SGTopic"), P(slide_label, "SGSmallNew")]], colWidths=[doc.width*0.76, doc.width*0.24])
        head.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),pale_blue),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LINEBEFORE",(0,0),(0,-1),4,blue),
        ]))
        story.extend([head, Spacer(1, 5)])
        if section.get("key_points"):
            story.extend([P("KEY IDEAS", "SGSubhead"), *bullets(section.get("key_points"))])
        if section.get("explanation"):
            story.extend([P("UNDERSTAND THE CONCEPT", "SGSubhead"), P(section["explanation"])])
        if section.get("why_it_matters"):
            story.append(panel("Why it matters", section["why_it_matters"], bg=navy_deep, accent=blue))
            story.append(Spacer(1, 5))
        if section.get("how_it_works"):
            story.append(P("HOW IT WORKS", "SGSubhead"))
            story.extend([P(f"{n}.  {step}", "SGSmallNew") for n,step in enumerate(section["how_it_works"],1)])
        if section.get("comparison"):
            story.append(panel("Compare and distinguish", section["comparison"], bg=navy, accent=colors.HexColor("#5CA5FF")))
            story.append(Spacer(1, 5))
        example = section.get("example") or {}
        if example:
            exbody = [str(example.get("scenario", ""))]
            exbody.extend(f"Step {n}. {step}" for n,step in enumerate(example.get("steps") or [],1))
            if example.get("takeaway"):
                exbody.append("Takeaway: " + str(example["takeaway"]))
            story.extend([P(example.get("title", "WORKED EXAMPLE"), "SGSubhead"), panel("Worked example", "<br/>".join(html.escape(x) for x in exbody), bg=colors.HexColor("#123B78"), accent=colors.HexColor("#58A8FF"))])
            story.append(Spacer(1, 5))
        for ex in section.get("additional_examples") or []:
            extra = [str(ex.get("scenario",""))] + [f"Step {n}. {v}" for n,v in enumerate(ex.get("steps") or [],1)]
            if ex.get("takeaway"):
                extra.append("Takeaway: " + str(ex["takeaway"]))
            story.append(panel(ex.get("title", "Additional example"), "<br/>".join(html.escape(x) for x in extra), bg=navy, accent=colors.HexColor("#58A8FF")))
            story.append(Spacer(1, 5))
        if section.get("common_mistakes"):
            story.extend([P("COMMON MISTAKES TO AVOID", "SGSubhead"), *bullets(section["common_mistakes"])])
        if section.get("brainhack"):
            # Text-drawn lightbulb marker remains compatible with standard PDF fonts.
            story.append(panel("BRAINHACK", section["brainhack"], bg=gold, label_color="SGGoldBody", body_style="SGGoldBody", accent=colors.HexColor("#E0A800")))
            story.append(Spacer(1, 5))
        terms = section.get("important_terms") or []
        if terms:
            story.append(P("IMPORTANT TERMS", "SGSubhead"))
            term_rows = [[P("<b>TERM</b>", "SGSmallNew"), P("<b>MEANING</b>", "SGSmallNew")]]
            term_rows.extend([[P(t.get("term",""),"SGSmallNew"),P(t.get("meaning",""),"SGSmallNew")] for t in terms])
            term_table = Table(term_rows, colWidths=[doc.width*0.29,doc.width*0.71], repeatRows=1)
            term_table.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),pale_blue),("BACKGROUND",(0,1),(-1,-1),colors.white),
                ("BOX",(0,0),(-1,-1),0.5,line),("INNERGRID",(0,0),(-1,-1),0.35,line),
                ("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),6),
                ("RIGHTPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),5),
                ("BOTTOMPADDING",(0,0),(-1,-1),5),
            ]))
            story.append(term_table)
        check = section.get("self_check") or {}
        if check.get("question"):
            check_text = html.escape(str(check["question"]))
            if check.get("answer"):
                check_text += "<br/><b>Answer:</b> " + html.escape(str(check["answer"]))
            story.append(panel("Quick self-check", check_text, bg=navy_deep, accent=colors.HexColor("#58A8FF")))
            story.append(Spacer(1, 5))
        if section.get("exam_focus"):
            story.extend([P("EXAM FOCUS", "SGSubhead"), *bullets(section["exam_focus"])])
        story.append(Spacer(1, 9))

    for heading, key in [("QUICK RECAP", "quick_recap"), ("EXAM SMART: REVISION TIPS", "exam_tips")]:
        values = guide.get(key) or []
        if values:
            story.extend([P(heading, "SGTopic"), *bullets(values)])
    if guide.get("coverage_note"):
        story.append(panel("Coverage and source notes", guide["coverage_note"], bg=navy, accent=colors.HexColor("#8AA9D6")))
    
    def page_background(canvas, document):
        canvas.saveState()
        canvas.setFillColor(page_blue)
        canvas.rect(0, 0, A4[0], A4[1], fill=1, stroke=0)
        canvas.setFillColor(navy_deep)
        canvas.rect(0, A4[1]-5*mm, A4[0], 5*mm, fill=1, stroke=0)
        canvas.setStrokeColor(line)
        canvas.setLineWidth(0.6)
        canvas.line(18*mm, 13*mm, A4[0]-18*mm, 13*mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(muted)
        canvas.drawRightString(A4[0]-18*mm, 8*mm, f"PAGE {document.page}")
        canvas.restoreState()

    if not story:
        story.append(P("No topic content was returned for this lecture.", "SGBodyNew"))
    doc.build(story, onFirstPage=page_background, onLaterPages=page_background)
    return buffer.getvalue()


@app.post("/api/study-guide-pdf")
async def study_guide_pdf(request: PdfRequest) -> StreamingResponse:
    pdf_started = time.perf_counter()
    content = await asyncio.to_thread(make_pdf, request.guide)
    LOGGER.info("pdf_created bytes=%s sections=%s duration_ms=%.1f", len(content), len(request.guide.get("sections") or []), (time.perf_counter() - pdf_started) * 1000)
    filename = re.sub(r"[^A-Za-z0-9_-]+", "_", str(request.guide.get("title", "StudentsGUIDE_Study_Guide"))).strip("_")[:70] or "StudentsGUIDE_Study_Guide"
    return StreamingResponse(
        io.BytesIO(content), media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}.pdf"'}
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=int(os.getenv("PORT", "8000")), reload=False)

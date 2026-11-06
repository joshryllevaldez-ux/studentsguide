from __future__ import annotations

import asyncio
import base64
import html
import io
import json
import os
import re
import secrets
import threading
import time
from collections import defaultdict, deque
from datetime import datetime, timezone
from typing import Any

import httpx
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, PlainTextResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from pptx import Presentation
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
APP_ACCESS_PASSWORD = os.getenv("APP_ACCESS_PASSWORD", "").strip()
IP_AI_GENERATION_LIMIT_PER_HOUR = int(os.getenv("IP_AI_GENERATION_LIMIT_PER_HOUR", "25"))
DAILY_AI_GENERATION_LIMIT = int(os.getenv("DAILY_AI_GENERATION_LIMIT", "100"))
AI_CONCURRENCY_LIMIT = max(1, int(os.getenv("AI_CONCURRENCY_LIMIT", "3")))
_AI_SEMAPHORE = asyncio.Semaphore(AI_CONCURRENCY_LIMIT)
_AI_LIMIT_LOCK = threading.Lock()
_AI_REQUESTS_BY_IP: dict[str, deque[float]] = defaultdict(deque)
_AI_DAILY_DATE = datetime.now(timezone.utc).date()
_AI_DAILY_COUNT = 0
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
if GEMINI_API_KEY.lower() in {"your_api_key_here", "replace_me", "changeme"}:
    GEMINI_API_KEY = ""
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
GEMINI_BASE_URL = "https://generativelanguage.googleapis.com/v1beta"

APP_VERSION = "0.3.1"
PROMPT_VERSION = "0.4"
app = FastAPI(title="StudentsGUIDE", version=APP_VERSION)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
app.mount("/assets", StaticFiles(directory=ASSET_DIR), name="assets")


def _basic_auth_is_valid(header: str | None) -> bool:
    """Validate the optional shared classroom password without storing credentials."""
    if not APP_ACCESS_PASSWORD:
        return True
    if not header or not header.lower().startswith("basic "):
        return False
    try:
        decoded = base64.b64decode(header.split(None, 1)[1], validate=True).decode("utf-8")
        _username, separator, supplied_password = decoded.partition(":")
        return bool(separator) and secrets.compare_digest(supplied_password, APP_ACCESS_PASSWORD)
    except (ValueError, UnicodeDecodeError, IndexError):
        return False


@app.middleware("http")
async def shared_access_and_security_headers(request: Request, call_next):
    # Optional shared class password. Configure APP_ACCESS_PASSWORD in the hosting dashboard.
    if request.url.path != "/api/status" and not _basic_auth_is_valid(request.headers.get("authorization")):
        return PlainTextResponse(
            "StudentsGUIDE class access is restricted. Enter the class access password in the browser prompt.",
            status_code=401,
            headers={"WWW-Authenticate": 'Basic realm="StudentsGUIDE Classroom", charset="UTF-8"'},
        )

    content_length = request.headers.get("content-length")
    if content_length:
        try:
            if int(content_length) > MAX_REQUEST_BYTES:
                return PlainTextResponse("Request is too large.", status_code=413)
        except ValueError:
            return PlainTextResponse("Invalid request size header.", status_code=400)

    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    return response


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


def extract_pptx_text(data: bytes) -> tuple[str, int]:
    try:
        prs = Presentation(io.BytesIO(data))
    except Exception as exc:
        raise HTTPException(status_code=400, detail="That file could not be opened as a PowerPoint presentation. Please upload a valid .pptx file.") from exc

    slide_blocks: list[str] = []
    for index, slide in enumerate(prs.slides, start=1):
        lines: list[str] = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text = "\n".join(p.text for p in shape.text_frame.paragraphs).strip()
                if text:
                    lines.append(text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        lines.append(row_text)
        if lines:
            slide_blocks.append(f"SLIDE {index}\n" + "\n".join(lines))

    text = "\n\n".join(slide_blocks).strip()
    if len(text) < 80:
        raise HTTPException(status_code=422, detail="I could not find enough readable text in this PowerPoint. Check that the slides contain selectable text rather than only screenshots.")
    if len(text) > MAX_SOURCE_CHARS:
        text = text[:MAX_SOURCE_CHARS] + "\n\n[Content shortened to fit the processing limit.]"
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


async def ask_ai(instructions: str, user_input: str, max_output_tokens: int = 6500) -> dict[str, Any]:
    if not GEMINI_API_KEY:
        raise HTTPException(
            status_code=503,
            detail="Gemini AI is not connected. The administrator must configure GEMINI_API_KEY on the hosting server.",
        )

    body = {
        "systemInstruction": {"parts": [{"text": instructions}]},
        "contents": [{
            "role": "user",
            "parts": [{
                "text": (
                    "Return only one valid JSON object. Do not use Markdown code fences. "
                    "The response must be JSON.\n\n" + user_input
                )
            }]
        }],
        "generationConfig": {
            "responseMimeType": "application/json",
            "maxOutputTokens": max_output_tokens,
        },
    }
    headers = {
        "x-goog-api-key": GEMINI_API_KEY,
        "Content-Type": "application/json",
    }

    try:
        async with _AI_SEMAPHORE:
            async with httpx.AsyncClient(
                timeout=httpx.Timeout(180.0, connect=20.0)
            ) as client:
                response = await client.post(
                    f"{GEMINI_BASE_URL}/models/{GEMINI_MODEL}:generateContent",
                    headers=headers,
                    json=body,
                )
    except httpx.TimeoutException as exc:
        raise HTTPException(
            status_code=504,
            detail="Gemini took too long to respond. Try a smaller PowerPoint or quiz.",
        ) from exc
    except httpx.HTTPError as exc:
        raise HTTPException(
            status_code=502,
            detail="StudentsGUIDE could not reach Gemini. Check the server connection.",
        ) from exc

    if response.status_code >= 400:
        try:
            error = response.json().get("error", {})
            detail = error.get("message", "Gemini rejected the request.")
        except Exception:
            detail = "Gemini rejected the request."
        raise HTTPException(
            status_code=429 if response.status_code == 429 else 502,
            detail=f"Gemini service error: {detail}",
        )

    payload = response.json()
    candidates = payload.get("candidates", [])
    if not candidates:
        raise HTTPException(
            status_code=502,
            detail="Gemini returned no response. Try a smaller lecture or quiz.",
        )

    candidate = candidates[0]
    if candidate.get("finishReason") == "MAX_TOKENS":
        raise HTTPException(
            status_code=502,
            detail="Gemini reached its response-length limit. Try a smaller lecture or quiz mode.",
        )

    parts = candidate.get("content", {}).get("parts", [])
    output = "\n".join(
        part.get("text", "")
        for part in parts
        if isinstance(part, dict) and part.get("text")
    ).strip()
    if not output:
        raise HTTPException(
            status_code=502,
            detail="Gemini returned an empty response. Please try again.",
        )
    return parse_json_text(output)

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
- Organize the guide into as many sections as the lecture needs (normally 5-14; more if necessary). Do not force 3-8 sections if that would merge distinct examinable topics.
- Each section should teach a coherent topic. Use clear headings and preserve the lecture's logical order.
- Explain concepts in plain, student-friendly language first, then retain the correct technical terminology.
- For each important concept, explain: what it is, how it works, why it matters, and where/how it is used when the source provides that information.
- Break processes, workflows, formulas, and mechanisms into numbered steps. Explain what happens at each step and why.
- Explain comparisons explicitly: show what is different, what is similar, and why the distinction matters. Do not invent comparison facts absent from the lecture.
- Include worked examples grounded in the lecture. Give at least one example per major section; give two or more when a section contains several distinct difficult concepts and the source supports them. Make examples concrete and explanatory, not merely restatements.
- For calculations or formulas, preserve exact source notation and values, define variables when provided, show each calculation step, and explain the result. If the source contains no calculation, do not invent one.
- Include common mistakes or misconceptions only when they can be derived safely from the source; label them as study cautions, not as claims made by the lecturer.
- Brainhacks must be genuinely useful retrieval cues: mnemonic, keyword association, analogy, sequence cue, or memory formula. Explain what each part of the cue means. Never use a catchy cue that distorts the facts.
- Add a quick self-check question and answer to each section to encourage active recall.
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
Return a complete guide. Do not omit a topic merely because it seems less important. Prioritize source fidelity, complete coverage, deep understanding, and student learning over a short PDF."""


QUIZ_INSTRUCTIONS = """You are StudentsGUIDE's assessment generator. Generate questions ONLY from the lecture content provided. Do not test unsupported outside knowledge. Keep answer keys and explanations accurate to the source. Return ONLY valid JSON, without markdown fences, in this shape:
{"quiz_title":"...","mode":"...","instructions":"...","questions":[{"id":1,"type":"mcq|true_false|scenario","topic":"...","context":"optional scenario context, otherwise empty string","question":"...","options":[{"label":"A","text":"..."},{"label":"B","text":"..."},{"label":"C","text":"..."},{"label":"D","text":"..."}],"correct_answer":"A|B|C|D|True|False","explanation":"complete explanation in clear language","example":"a short illustrative example; make the first question's example particularly detailed and step-by-step","brainhack":"a concise memory cue that preserves accuracy","source_basis":"briefly identify which lecture idea supports the answer"}]}
For true_false questions, use exactly two options labelled True and False, and set correct_answer to the word True or False. For mcq and scenario questions, use four plausible options and set correct_answer to the correct option label. For scenario questions, give a realistic short situation that requires applying the lecture concept. Avoid trick questions and avoid repeating the same concept. The first question's explanation and example must be especially thorough. All other explanations must still be complete and teach why the answer is correct. Return exactly the requested number of questions."""


@app.get("/")
async def home() -> FileResponse:
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


@app.get("/api/status")
async def status() -> dict[str, Any]:
    return {
        "ok": True,
        "app_version": APP_VERSION,
        "ai_configured": bool(GEMINI_API_KEY),
        "model": GEMINI_MODEL if GEMINI_API_KEY else None,
        "max_upload_mb": MAX_UPLOAD_MB,
        "supported_file": ".pptx",
        "shared_access_enabled": bool(APP_ACCESS_PASSWORD),
        "max_ai_generations_per_day": DAILY_AI_GENERATION_LIMIT,
        "ip_ai_generations_per_hour": IP_AI_GENERATION_LIMIT_PER_HOUR,
    }


@app.post("/api/generate-study-guide")
async def generate_study_guide(request: Request, file: UploadFile = File(...)) -> dict[str, Any]:
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=503, detail="AI generation is not connected yet. The StudentsGUIDE administrator must configure the API key on the hosting server.")
    enforce_ai_generation_quota(request)
    filename = file.filename or "lecture.pptx"
    if not filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Please upload a PowerPoint .pptx file.")
    data = await file.read()
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail=f"The file is larger than {MAX_UPLOAD_MB} MB. Please upload a smaller PowerPoint.")
    if not data:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")

    source_text, slide_count = extract_pptx_text(data)
    prompt = (
        f"Create a study guide from this lecture PowerPoint. Original file: {filename}. Slides detected: {slide_count}.\n\n"
        "LECTURE CONTENT (organized by slide):\n" + source_text
    )
    guide = await ask_ai(STUDY_GUIDE_INSTRUCTIONS, prompt, max_output_tokens=12000)
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
    enforce_ai_generation_quota(request)
    mode = payload.mode.upper().strip()
    if mode not in MODE_SPECS:
        raise HTTPException(status_code=400, detail="Choose one of the five available quiz modes.")
    count, mix_instruction = MODE_SPECS[mode]
    prompt = (
        f"Lecture title: {payload.lecture_title}\nSelected mode: {mode}\n"
        f"Requirement: {mix_instruction}\nExact number of questions: {count}.\n\n"
        "LECTURE CONTENT:\n" + payload.source_text[:MAX_SOURCE_CHARS]
    )
    quiz = await ask_ai(QUIZ_INSTRUCTIONS, prompt, max_output_tokens=11000 if count >= 40 else 7500)
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
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=19*mm, leftMargin=19*mm,
        topMargin=18*mm, bottomMargin=18*mm,
        title=str(guide.get("title", "StudentsGUIDE Study Guide")),
        author="StudentsGUIDE.com",
    )
    navy = colors.HexColor("#142F69")
    blue = colors.HexColor("#087CEB")
    pale = colors.HexColor("#EAF3FF")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SGTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=23, leading=28, textColor=navy, alignment=TA_LEFT, spaceAfter=8))
    styles.add(ParagraphStyle(name="SGSub", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=13, textColor=colors.HexColor("#5E6E86"), spaceAfter=10))
    styles.add(ParagraphStyle(name="SGH1", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=15, leading=19, textColor=navy, spaceBefore=14, spaceAfter=7, keepWithNext=True))
    styles.add(ParagraphStyle(name="SGH2", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=10.5, leading=14, textColor=blue, spaceBefore=7, spaceAfter=3, keepWithNext=True))
    styles.add(ParagraphStyle(name="SGBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=9.3, leading=14, textColor=colors.HexColor("#263348"), spaceAfter=6))
    styles.add(ParagraphStyle(name="SGSmall", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=12, textColor=colors.HexColor("#34445D"), spaceAfter=4))

    def P(value: Any, style: str = "SGBody") -> Paragraph:
        return Paragraph(html.escape(str(value or "")).replace("\n", "<br/>"), styles[style])

    story: list[Any] = []
    logo_path = os.path.join(ASSET_DIR, "studentsguide-brand-pdf.jpg")
    if os.path.exists(logo_path):
        img = RLImage(logo_path, width=172*mm, height=57.3*mm)
        img.hAlign = "CENTER"
        story.extend([img, Spacer(1, 6)])
    story.append(P(guide.get("title", "Study Guide"), "SGTitle"))
    meta = f"{guide.get('course', 'Course not specified')}  •  {guide.get('source_filename', 'Uploaded lecture')}  •  {guide.get('slide_count', '')} slides"
    story.append(P(meta, "SGSub"))
    story.append(P("STUDY GUIDE OVERVIEW", "SGH2"))
    story.append(P(guide.get("overview", "A structured review of the uploaded lecture.")))

    objectives = guide.get("learning_objectives") or []
    if objectives:
        story.append(P("What you should understand", "SGH2"))
        for item in objectives:
            story.append(P("• " + str(item), "SGSmall"))

    for index, section in enumerate(guide.get("sections") or [], start=1):
        block: list[Any] = [P(f"{index}. {section.get('heading', 'Key topic')}", "SGH1")]
        if section.get("key_points"):
            block.append(P("Key points", "SGH2"))
            for point in section.get("key_points", []):
                block.append(P("• " + str(point), "SGSmall"))
        if section.get("explanation"):
            block.extend([P("Simplified explanation", "SGH2"), P(section["explanation"])])
        example = section.get("example") or {}
        if example:
            block.append(P(example.get("title", "Example"), "SGH2"))
            if example.get("scenario"):
                block.append(P(example["scenario"]))
            for step_no, step in enumerate(example.get("steps") or [], start=1):
                block.append(P(f"{step_no}. {step}", "SGSmall"))
            if example.get("takeaway"):
                block.append(P("Takeaway: " + str(example["takeaway"]), "SGSmall"))
        if section.get("brainhack"):
            hack = Table([[P("BRAIN HACK", "SGH2")], [P(section["brainhack"], "SGSmall")]], colWidths=[doc.width])
            hack.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,-1), pale), ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#C8DDF7")),
                ("LEFTPADDING", (0,0), (-1,-1), 9), ("RIGHTPADDING", (0,0), (-1,-1), 9),
                ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ]))
            block.extend([Spacer(1, 3), hack, Spacer(1, 5)])
        terms = section.get("important_terms") or []
        if terms:
            block.append(P("Important terms", "SGH2"))
            for term in terms:
                block.append(P(f"{term.get('term', '')}: {term.get('meaning', '')}", "SGSmall"))
        story.append(KeepTogether(block))

    for heading, key in [("Quick recap", "quick_recap"), ("Exam tips", "exam_tips")]:
        values = guide.get(key) or []
        if values:
            story.append(P(heading, "SGH1"))
            for value in values:
                story.append(P("• " + str(value), "SGSmall"))
    if guide.get("coverage_note"):
        story.append(P("Content coverage note", "SGH2"))
        story.append(P(guide["coverage_note"], "SGSmall"))
    story.append(Spacer(1, 10))
    story.append(P("Generated by StudentsGUIDE.com • Struggles less, study smart.", "SGSub"))

    def footer(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D9E5F3"))
        canvas.line(19*mm, 14*mm, A4[0]-19*mm, 14*mm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#71809A"))
        canvas.drawString(19*mm, 9*mm, "StudentsGUIDE.com")
        canvas.drawRightString(A4[0]-19*mm, 9*mm, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()


@app.post("/api/study-guide-pdf")
async def study_guide_pdf(request: PdfRequest) -> StreamingResponse:
    content = make_pdf(request.guide)
    filename = re.sub(r"[^A-Za-z0-9_-]+", "_", str(request.guide.get("title", "StudentsGUIDE_Study_Guide"))).strip("_")[:70] or "StudentsGUIDE_Study_Guide"
    return StreamingResponse(
        io.BytesIO(content), media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}.pdf"'}
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=int(os.getenv("PORT", "8000")), reload=False)

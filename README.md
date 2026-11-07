# StudentsGUIDE.com — AI Learning Platform

**Tagline:** Struggles less, study smart.  
**Development version:** StudentsGUIDE 2.0 (development branch)  
**Production branch:** main  
**Development branch:** studentsguide-2.0-dev

StudentsGUIDE converts lecture PowerPoint files into structured study guides, downloadable PDFs, and lecture-grounded practice quizzes. The application uses FastAPI, a browser-based frontend, ReportLab, python-pptx, and a server-side generative AI provider.

## Current features

### PowerPoint processing
- Accepts .pptx uploads with configurable size limits.
- Extracts selectable text and table cells slide by slide.
- Retains slide numbering in the extracted source.
- Explicitly flags slides with no selectable text and slides containing embedded images that have not been OCR-analyzed.
- Reports truncation rather than implying that a truncated lecture was fully covered.

### AI study guide
- Produces topic sections, explanations, examples, memory cues, key terms, and revision points.
- Includes source slide references where available.
- Instructs the model to disclose unclear or unprocessed source content.
- Stores the source text with the generated guide for lecture-based quiz generation.

### PDF export
- Begins directly with lecture topics rather than a promotional cover card.
- Uses a restrained blue academic design, description panels, example panels, and distinct Brain Hack callouts.
- Includes page numbering and checks for valid PDF output in automated tests.

### Interactive quizzes
- Supports MCQ 20, MCQ 10, True/False 10, Scenario 15, and Mixed 40.
- Includes answer selection, instant feedback, explanations, examples, Brain Hacks, review flags, navigation, submission, scoring, and incorrect-answer review.

### Learning progress
- Stores up to 20 real quiz attempts in browser local storage.
- Displays recent attempts, question-weighted overall accuracy, and weak topics from the latest attempt.
- Does not store sample-preview scores.
- Progress is local to the current browser/device. It is not an account, cloud sync, or cross-device history.

## AI provider configuration

The hosted Render configuration declares OpenAI environment variables. The backend now uses OpenAI as the preferred provider and retains Gemini as a backwards-compatible local fallback.

For local use:
1. Copy `.env.example` to `.env`.
2. Add your own API key to `OPENAI_API_KEY`.
3. Keep the key private. Never commit `.env` or expose API keys in browser code.
4. Install dependencies with `python -m pip install -r requirements.txt`.
5. Run the FastAPI app with `uvicorn app:app --reload`.

The model name is configured by `OPENAI_MODEL`; use a model identifier enabled for the API account. API usage is billed separately from ChatGPT subscription usage.

For Render, set `OPENAI_API_KEY` in the service's environment settings. `render.yaml` declares this variable as a secret that must be supplied in the dashboard.

## Development and tests

Install development dependencies and run:
```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

The tests cover PowerPoint extraction, slide-gap warnings, JSON parsing, and PDF creation. They do not prove the factual accuracy of live model output or replace browser-based end-to-end testing.

## Important limitations and planned work

- Embedded images are detected but their contents are not yet OCR- or vision-analyzed. Image-only diagrams, charts, and screenshots must be reviewed against the original PowerPoint.
- AI-generated study material and answer keys can still contain errors. Review them against the lecture before relying on them for examinations.
- Learning progress currently uses local browser storage, not a persistent multi-user database.
- There is no student account system, cloud study library, cross-device synchronization, or cloud file retention workflow.
- An AI tutoring chat, spaced-repetition scheduler, timed exam simulation, robust adaptive question selection, and full OCR/diagram interpretation remain future development items.
- Before broad public release, add a privacy-reviewed account and storage architecture, durable rate limiting, file retention/deletion controls, monitoring, abuse protection, and a security review.

## Deployment

The Render service is configured for automatic deployment from commits to `main`. The 2.0 development work is isolated on `studentsguide-2.0-dev` and does not replace the live production branch until it has been reviewed and merged.

Do not describe a feature as tested or deployed unless the corresponding test or deployment has actually been verified.

## Student pricing and payment policy — non-negotiable

**Students must be able to use the core StudentsGUIDE learning workflow at AED 0.** Uploading or dragging and dropping a PowerPoint, generating a study guide or PDF, and using the included practice-quiz features must not trigger a student payment.

- No student subscriptions, paywalls, credit-card collection, per-upload fees, paid credits, or payment prompts for the core learning workflow.
- Do not add Stripe, PayPal, or another payment processor to the student experience.
- Do not add hidden charges or route students to a payment page when they upload a file.
- AI/API, hosting, and storage charges (if any) are operating costs paid by the platform owner. They must never be passed to students through an upload fee or surprise charge.
- Protect operating costs with reasonable, clearly disclosed fair-use limits, upload-size limits, request throttling, caching, and monitoring. A limit must not become a payment gate.
- If the platform cannot afford a request, it should explain that the service is temporarily unavailable or that a fair-use limit was reached; it must not ask the student to pay.

Any future proposal that changes this policy requires the owner's explicit approval before implementation. This policy applies to development and production.

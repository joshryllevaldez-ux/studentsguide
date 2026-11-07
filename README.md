# StudentsGUIDE.com — AI Application

**Tagline:** Struggles less, study smart.  
**Version:** 0.4.3  
**Project:** AIT111 Artificial Intelligence, Fall 2026–2027

StudentsGUIDE is a locally run educational AI application. It takes a student's lecture PowerPoint presentation, extracts the readable slide text, and uses a Generative AI model to create a structured study guide and an interactive quiz. It runs on the student's Windows computer and opens in the default browser at a local address. Public website hosting or a purchased domain is not required for the application demonstration.

## Main features

### PowerPoint to study guide

- Drag-and-drop or browse to upload a `.pptx` file (maximum 50 MB).
- Extracts selectable text and table text slide-by-slide.
- Generates a structured study guide using the lecture content:
  - Overview and learning objectives
  - Topic sections and key points
  - Simplified explanations
  - Worked examples, with the first example especially detailed
  - Brain hacks and memory cues
  - Important terms, recap and exam tips
- Automatically creates and downloads a branded PDF.

### Interactive quiz generator

The student can choose one of the following modes:

| Mode | Question count and composition |
|---|---:|
| MCQ 20 | 20 multiple-choice questions |
| MCQ 10 | 10 multiple-choice questions |
| True or False 10 | 10 true/false questions |
| Scenario Questions 15 | 15 applied scenario questions |
| Mixed 40 | 16 MCQ, 12 true/false and 12 scenario questions |

Quiz controls include answer selection, See Answer, full explanation, example, brain hack, Mark to Review, Previous, Next, question navigation, submission, score calculation, incorrect-answer review and retaking the quiz.

## Run the application on Windows

1. Download and extract the `StudentsGUIDE_AI_Application_Windows.zip` file.
2. Install Python 3.10 or newer from [python.org](https://www.python.org/downloads/). During installation, select **Add Python to PATH**.
3. Open the extracted `StudentsGUIDE` folder.
4. Double-click `start_studentsguide.bat`.
5. On the first run, the launcher creates a Python virtual environment and installs the required packages. Internet access is needed for this first setup.
6. Your default browser opens when the application server is ready. The application runs locally at `http://127.0.0.1:<port>`.
7. Keep the black launcher window open while using the application. Press `CTRL+C` in that window to stop it.

The launcher chooses an available local port, starting at 8000. It does not publish the app to the internet.

## Connect live AI generation

StudentsGUIDE uses the Google Gemini Developer API. Create an API key in [Google AI Studio](https://aistudio.google.com/apikey).

### Render deployment
1. Open the StudentsGUIDE service in Render.
2. Open **Environment**.
3. Add `GEMINI_API_KEY` as a secret environment variable and paste the key.
4. Set `GEMINI_MODEL` to `gemini-3.8-flash` and `GEMINI_FAST_MODEL` to `gemini-3.5-flash-lite` if they are not already set.
5. Save changes and wait until the deployment is Live.

### Local Windows run
1. Copy `.env.example` to `.env` if the launcher has not already done so.
2. Open `.env` in Notepad and set `GEMINI_API_KEY=your_key_here`.
3. Keep `GEMINI_MODEL=gemini-3.8-flash` and `GEMINI_FAST_MODEL=gemini-3.5-flash-lite`.
4. Save the file and run `start_studentsguide.bat` again.

The API key is read only by the Python server. Never place it in frontend JavaScript or commit a populated `.env` file to GitHub. Gemini API free-tier use has quotas and rate limits; check Google's current [Gemini API pricing](https://ai.google.dev/gemini-api/docs/pricing) and [rate limits](https://ai.google.dev/gemini-api/docs/rate-limits). A free tier is not an unlimited service.
## Technology

- Frontend: HTML, CSS and JavaScript
- Backend: Python and FastAPI
- PowerPoint extraction: `python-pptx`
- Generative AI: Gemini 3.8 Flash for study guides, scenario quizzes and Mixed 40; Gemini 3.5 Flash-Lite for MCQ and True/False quizzes
- Structured JSON output: Gemini `responseMimeType: application/json`
- Study guide PDF: ReportLab
- Local launcher: Python + Uvicorn + Windows batch launcher

Model capabilities, supported thinking levels and output limits are documented in the [Gemini 3.8 Flash model reference](https://ai.google.dev/gemini-api/docs/models/gemini-3.8-flash) and [Gemini thinking guide](https://ai.google.dev/gemini-api/docs/thinking).

## Important limits

- The current PowerPoint processor extracts text from text boxes, grouped shapes, tables, charts and speaker notes. It does not run OCR on text that exists only inside screenshots or images.
- Charts, diagrams and image-based slides may need to be checked against the original PowerPoint.
- To avoid silently omitting source material, a deck whose extracted text exceeds the 55,000-character processing limit is rejected with instructions to split the PowerPoint into smaller parts.
- The AI is instructed to use only the supplied lecture content, but generated material can still contain errors. Review study guides and answer keys against the source slides before using them for assessment preparation.
- The sample experience is illustrative preview content. It is not an AI-generated output and must not be presented as such in the AIT111 report.
- The application does not currently include student logins, cloud storage, multiple-user accounts, analytics or public hosting.
- A public release would require authentication, rate limits, usage quotas, file retention and deletion rules, HTTPS, monitoring, secure secrets management and a security review.

## Testing and development evidence

The restored `main` branch currently does not contain an automated test suite. Do not report tests as passed unless a test suite is added and run against this version.

Real API calls and end-to-end testing with an actual course lecture must be completed after the Gemini key is configured. Use `TESTING_CHECKLIST.md` to record real test dates, the uploaded course file, model ID, prompt version, generation time and any content issues. Backend timing logs are available in the Render service logs for HTTP requests, PPTX extraction, Gemini generation and PDF creation.
## Project files

- `app.py` — FastAPI backend, PowerPoint extraction, model requests and PDF creation
- `static/index.html` — StudentsGUIDE interface and interactive quiz behaviour
- `assets/` — official supplied brand banner and campus-inspired illustration
- `run.py` — local app launcher
- `start_studentsguide.bat` — Windows one-click launcher
- `.env.example` — API configuration template
- `PROMPT_ENGINEERING_LOG.md` — prompt development record
- `TESTING_CHECKLIST.md` — evaluation checklist

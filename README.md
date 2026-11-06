# StudentsGUIDE.com — AI Application

**Tagline:** Struggles less, study smart.  
**Version:** 0.2.0  
**Project:** AIT111 Artificial Intelligence, Fall 2026–2027

StudentsGUIDE is a locally run educational AI application. It takes a student's lecture PowerPoint presentation, extracts the readable slide text, and uses a Generative AI model to create a structured study guide and an interactive quiz. It runs on the student's Windows computer and opens in the default browser at a local address. Public website hosting or a purchased domain is not required for the application demonstration.

## Main features

### PowerPoint to study guide

- Drag-and-drop or browse to upload a `.pptx` file (maximum 20 MB).
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

The user must configure an OpenAI API key on the local server before using live AI generation. A ChatGPT subscription does not automatically supply API credits.

1. Start StudentsGUIDE once so it creates `.env` from `.env.example`.
2. Close StudentsGUIDE.
3. Open `.env` in Notepad.
4. Replace `your_api_key_here` after `OPENAI_API_KEY=` with the API key from the OpenAI developer platform.
5. Leave `OPENAI_MODEL=gpt-6.1-terra` unless you intentionally want to use another model available to your API account.
6. Save `.env` and run `start_studentsguide.bat` again.

The API key is read only by the Python server. It is not included in the browser interface or returned by the status endpoint. Keep `.env private and do not share the application folder with a configured key inside it.

API usage is billed separately from ChatGPT subscription usage. Set API billing limits before testing.

## Technology

- Frontend: HTML, CSS and JavaScript
- Backend: Python and FastAPI
- PowerPoint extraction: `python-pptx`
- Generative AI: OpenAI Responses API
- Structured JSON output: Responses API JSON mode
- Study guide PDF: ReportLab
- Local launcher: Python + Uvicorn + Windows batch launcher

The default API model ID, `gpt-6.1-terra`, is listed in OpenAI's API model documentation. See [OpenAI API models](https://developers.openai.com/api/docs/models) and [Structured Outputs / JSON mode](https://developers.openai.com/api/docs/guides/structured-outputs?api-mode=responses).

## Important limits

- The current PowerPoint processor extracts text and table cells. It does not run OCR on text that exists only inside screenshots or images.
- Charts, diagrams and image-based slides may need to be checked against the original PowerPoint.
- The AI is instructed to use only the supplied lecture content, but generated material can still contain errors. Review study guides and answer keys against the source slides before using them for assessment preparation.
- The sample experience is illustrative preview content. It is not an AI-generated output and must not be presented as such in the AIT111 report.
- The application does not currently include student logins, cloud storage, multiple-user accounts, analytics or public hosting.
- A public release would require authentication, rate limits, usage quotas, file retention and deletion rules, HTTPS, monitoring, secure secrets management and a security review.

## Testing and development evidence

- Automated tests are in `tests/test_app.py`.
- Run the tests with `python -m pip install -r requirements-dev.txt` followed by `python -m pytest -q`.
- Prompt versions and planned evidence are recorded in `PROMPT_ENGINEERING_LOG.md`.
- Real API calls and end-to-end testing with an actual course lecture must be completed after the API key is configured.
- Do not report pending tests as passed. Record the actual course file, model, prompt version and date for each real AI test.

## Project files

- `app.py` — FastAPI backend, PowerPoint extraction, model requests and PDF creation
- `static/index.html` — StudentsGUIDE interface and interactive quiz behaviour
- `assets/` — official supplied brand banner and campus-inspired illustration
- `run.py` — local app launcher
- `start_studentsguide.bat` — Windows one-click launcher
- `.env.example` — API configuration template
- `tests/` — automated backend tests
- `PROMPT_ENGINEERING_LOG.md` — prompt development record
- `TESTING_CHECKLIST.md` — evaluation checklist

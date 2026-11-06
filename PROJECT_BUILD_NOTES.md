# StudentsGUIDE — Build Notes and Current State

**Build date:** 6 November 2026  
**Application version:** 0.2.0  
**Project:** AIT111 Artificial Intelligence  
**Product:** StudentsGUIDE.com — “Struggles less, study smart.”

## What has been built

- Local, responsive AI application using the official user-supplied StudentsGUIDE banner.
- Custom campus-inspired vector illustration, clearly labelled as illustrative rather than official Ajman University artwork.
- Drag-and-drop `.pptx` upload with a 20 MB limit and file validation.
- FastAPI server extracts selectable slide text and table text slide-by-slide.
- Server-side OpenAI Responses API integration. The API key is read from `.env` and is never embedded in client-side JavaScript.
- JSON-mode Responses API output to make generated JSON easier to parse.
- AI study guide with overview, learning objectives, topic sections, key points, simplified explanations, examples, brain hacks, important terms, recap, exam tips and content coverage note.
- Branded PDF download with the supplied StudentsGUIDE banner, page numbering and optimized logo image.
- Quiz generation in five modes: MCQ 20, MCQ 10, True or False 10, Scenario Questions 15 and Mixed 40.
- Mixed 40 requires 16 MCQ, 12 True/False and 12 Scenario questions; the server validates this distribution.
- Quiz controls: answer selection, See Answer, detailed explanation, example, brain hack, Mark to Review, Previous, Next, question map, Submit, score, incorrect-answer review and retake.
- Branded rotating-logo loading screen after opening the quiz and while generating quizzes.
- Illustrative sample preview for exploring the design without an API key. It is explicitly labelled and is not live AI content.
- Local Windows launcher that creates a virtual environment, installs dependencies, selects a free local port and opens the browser when the server is ready.
- Prompt engineering record, technical specification, user setup instructions and testing checklist.

## Technical architecture

1. Browser sends a `.pptx` file to the local FastAPI backend.
2. `python-pptx` extracts text from slides and table cells.
3. Backend sends lecture text to the OpenAI Responses API using the configured model and server-side API key.
4. Responses API JSON mode returns structured content for the frontend.
5. Backend validates study-guide or quiz output.
6. ReportLab creates the branded PDF.
7. Browser handles quiz navigation, selected answers, marking, answer reveal, scoring and results review.

## Automated tests completed

Command: `python -m pytest -q`  
Result: **11 passed**.

Tests cover:

- App status and version endpoint.
- PowerPoint text extraction from a generated test presentation.
- Study-guide route with a mocked AI response.
- PDF output signature and media type.
- All five quiz modes with deterministic mocked model responses.
- Enforcement of the Mixed 40 distribution.
- Responses API URL, model ID, JSON-mode payload and server-side authorization header construction.

The JavaScript bundle passed `node --check` syntax validation.

## Not yet verified

- A live OpenAI API request has not been run because no API key was supplied or configured in the build environment.
- Full end-to-end generation with the student's actual course lecture has not been completed.
- Semantic accuracy of generated examples, brain hacks, questions and answer keys must be checked against the original lecture.
- The Windows launcher must be run on the student's Windows laptop to verify first-run installation and browser startup in that environment.
- Interactive browser automation was not completed in the build environment; browser controls should be manually checked on Windows.

## Current constraints

- PowerPoint text and table text are extracted. Text embedded only in images/screenshots is not OCR-processed.
- Charts, diagrams and image-based slides may require manual review.
- Model output can still be inaccurate even when prompted to stay within the lecture content.
- Local, single-computer prototype only. No student accounts, multi-user cloud storage, analytics, public hosting or domain registration is configured.
- Public release requires authentication, rate limits, usage quotas, file retention/deletion rules, HTTPS, monitoring, secure secret management and a security review.

## Current release package

`StudentsGUIDE_AI_Application_Windows.zip` contains the source code, assets, launcher, sample PowerPoint, automated tests and documentation. It is run locally and requires Python 3.10+ plus an OpenAI API key for live generation.

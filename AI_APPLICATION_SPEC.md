# StudentsGUIDE AI Application — Product Specification

## Purpose

StudentsGUIDE turns readable content from a lecture PowerPoint into two linked learning outputs: a downloadable study guide and an interactive quiz.

## User flow

1. Student launches the local application.
2. Student selects or drags a PowerPoint `.pptx` file into the upload area.
3. The server checks file type and size, extracts slide text and table contents, and submits the lecture text to the configured language model.
4. The model generates a study guide containing an overview, learning objectives, topic sections, key points, simplified explanations, examples, brain hacks, important terms, quick recap and exam tips.
5. The server produces a branded PDF. The student can also download the guide later from the guide view.
6. Student selects a quiz mode: MCQ 20, MCQ 10, True or False 10, Scenario Questions 15, or Mixed 40.
7. The server requests lecture-grounded questions, answer keys, detailed explanations, examples and brain hacks.
8. Student answers, marks questions for review, navigates, reveals explanations and submits the quiz.
9. The application calculates the score and displays incorrect-answer review.

## AI model behaviour

- The model receives extracted lecture text, not the entire original PowerPoint file.
- The study-guide prompt instructs the model to rely only on the uploaded lecture and disclose unclear or missing content.
- The quiz prompt instructs the model to avoid testing unsupported outside knowledge.
- The first worked example and first quiz explanation/example are requested to be especially detailed.
- Mixed 40 has an enforced distribution of 16 MCQ, 12 True/False and 12 Scenario questions.
- The backend validates the returned question objects before displaying them.

## Architecture

`PPTX file → FastAPI upload endpoint → python-pptx text extraction → OpenAI Responses API → structured JSON → browser study guide / quiz UI → ReportLab PDF`

The browser communicates only with the local FastAPI server. The API key is kept in the server-side `.env` file and is not embedded in client-side JavaScript.

## Current scope and known limitations

- Local use on one computer; no public cloud deployment.
- Selectable PowerPoint text and table text are extracted.
- Text embedded only in slide images is not OCR-processed.
- No student accounts, persistence across sessions, cloud storage or history.
- Real AI generation requires an API key and separate API billing.
- Model-generated content requires human review against the original lecture.

## AIT111 demonstration plan

Demonstrate a real lecture file, show extracted/generated guide content, open the resulting PDF, create one quiz mode, reveal an answer explanation and brain hack, mark a question for review, navigate, submit and display the results. Keep a screen recording as a backup. Do not use sample preview material as evidence of live generation.

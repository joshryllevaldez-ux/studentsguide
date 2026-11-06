# StudentsGUIDE Prompt Engineering Evidence Log

The AIT111 final report requires the original prompt, refined prompts, reasons for changes and evidence of the prompt process. This file records the prompts and engineering choices actually present in the application. Add screenshots and findings only after running real AI generations.

## Prompt Version 0.1 — Initial structured generation design

### Study guide generation prompt

The original implementation prompt asked the model to use lecture text only and return a structured guide with:

- Title, course, overview and learning objectives.
- Three to eight topic sections, sized to the amount of lecture content.
- Key points and simplified explanations.
- A worked example with steps and takeaway.
- A brain hack, important terms, quick recap and exam tips.
- A content coverage note when slide content is unclear or incomplete.

### Quiz generation prompt

The original implementation prompt asked the model to create lecture-grounded MCQ, true/false and scenario questions with options, answer key, explanation, example, brain hack and source basis.

## Prompt Version 0.2 — Refinement for the working application

The current prompt templates are in `app.py` as `STUDY_GUIDE_INSTRUCTIONS` and `QUIZ_INSTRUCTIONS`.

### Changes and reasons

| Refinement | Implemented behaviour | Reason |
|---|---|---|
| Content grounding | The prompts instruct the model to use only the supplied lecture and disclose unclear/missing information. | Reduce unsupported explanations and questions. |
| Teaching detail | The first worked example and first quiz explanation/example are requested to be especially detailed and step-by-step. | Demonstrate a clear, thorough teaching example rather than only short generated answers. |
| Structured response | The Responses API request uses JSON mode, and the server validates the returned object. | Make the generated content easier for the app to render and handle. |
| Mode-specific constraints | Each quiz mode requests an exact number and type of questions. Mixed 40 requires 16 MCQ, 12 True/False and 12 Scenario questions. | Match the five quiz modes selected for StudentsGUIDE. |
| Quiz quality | Each question includes an answer explanation, example, brainhack and source basis. | Make the quiz a learning tool, not just a list of questions. |

## Current prompt templates

The complete current prompt text is kept in the source code so its exact wording can be captured and included in the final report:

- `STUDY_GUIDE_INSTRUCTIONS`
- `QUIZ_INSTRUCTIONS`

## Evidence still to collect

After an API key is configured and the first real lecture is processed, save:

1. A screenshot of the exact prompt version used.
2. The lecture file name, date, model ID and prompt version.
3. A sample generated study guide and quiz output.
4. Any issue found during review against the slides.
5. The refined prompt and explanation of why it was changed.
6. A before/after example showing the observed improvement.

Do not invent test results or describe the sample preview as AI output. The prompt iteration record above documents engineering choices; model-output improvement must be evaluated with actual tests.

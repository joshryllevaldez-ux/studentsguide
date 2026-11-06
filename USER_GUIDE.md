# StudentsGUIDE — Quick User Guide

## First-time setup

1. Extract the application ZIP.
2. Install Python 3.10 or newer.
3. Double-click `start_studentsguide.bat`.
4. Wait for the first-run package installation to finish. The browser opens when StudentsGUIDE is ready.
5. Keep the command window open during use.

## Enable actual AI generation

1. Close StudentsGUIDE.
2. Open `.env` in the extracted application folder.
3. Set `OPENAI_API_KEY=...` to your own API key.
4. Save the file and restart the app.
5. Check the AI connection status on the homepage.

The application uses the OpenAI Responses API and defaults to `gpt-6.1-terra`. API use may incur charges independently of a ChatGPT subscription.

## Generate a study guide

1. Drag a `.pptx` lecture into the upload area or select **browse files**.
2. Confirm that the file name and size appear.
3. Click **Generate my Study Guide**.
4. Wait for the AI to process the lecture.
5. Review the guide and save the automatically downloaded PDF. Use **Download Study Guide PDF** to download again.

The generated guide is intended to include a lecture overview, learning objectives, key points, simplified explanations, worked examples, brain hacks, important terms, a quick recap and exam tips.

## Create an interactive quiz

1. Click **Go to my Interactive Quiz**.
2. Wait for the branded loading animation.
3. Choose MCQ 20, MCQ 10, True or False 10, Scenario Questions 15 or Mixed 40.
4. Select an answer for each question.
5. Use **See answer & explanation** to reveal the correct answer, explanation, example and brain hack.
6. Use **Mark to review**, Previous, Next or the numbered question map.
7. Click **Submit quiz** to see the score and answer review.

## Important notes

- Upload only `.pptx` files up to 20 MB.
- Only selectable slide text and table text are extracted. Text inside screenshots is not OCR-processed.
- Check the generated material against the original lecture. AI can make mistakes.
- The sample preview is illustrative only; it is not generated from your lecture.
- Do not upload private or restricted course material unless you have permission.

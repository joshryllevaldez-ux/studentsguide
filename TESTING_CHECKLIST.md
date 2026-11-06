# StudentsGUIDE AI Application — Testing Checklist

## Automated tests completed in the build environment

`python -m pytest -q` — **11 passed**.

Coverage includes:

- Application status endpoint.
- PowerPoint slide text extraction.
- Study-guide generation route with a mocked model response.
- PDF generation route and PDF signature validation.
- All five quiz-generation modes with mocked model responses.
- Rejection of an incorrect Mixed 40 question distribution.
- OpenAI Responses API request construction, including JSON mode and server-side authorization header.

These are application/backend tests using deterministic mock responses. They do not prove the semantic accuracy of live AI-generated study materials.

## Manual testing required with a real lecture

| Test | Expected result | Status |
|---|---|---|
| Launch on Windows | Launcher installs dependencies and opens local application | Pending on user's Windows laptop |
| Upload a valid `.pptx` | File name appears and Generate button enables | Pending with user's lecture |
| Upload a non-PPTX file | Clear format warning | Frontend validation present; manual browser check pending |
| Upload a PPTX over 20 MB | Clear size warning | Frontend/backend validation present; manual browser check pending |
| Process a lecture with selectable text | Slide text is extracted and sent to the model | Extraction unit test passed; live AI pending |
| Process a mostly image-based deck | App explains that image-only text is not extracted | Backend behaviour present; manual check pending |
| Generate study guide with live AI | Overview, objectives, explanations, worked examples, brain hacks and recap appear | Requires API key and real lecture |
| Export study guide PDF | Branded PDF downloads and opens | PDF endpoint automated test passed; real guide PDF pending |
| Generate each quiz mode | Requested number and question types are returned | Route tests passed with mocked model; live AI pending |
| See Answer | Correct answer, explanation, example and brainhack appear | Browser interaction check pending |
| Mark to Review | Question is visibly flagged | Browser interaction check pending |
| Previous / Next | Navigation works and answers persist | Browser interaction check pending |
| Submit quiz | Score and answer review appear | Browser interaction check pending |
| Verify content grounding | Questions and explanations match the lecture | Manual academic review required |

Record actual date, lecture name, model ID, prompt version and findings for every live AI test before using results in the final AIT111 report.

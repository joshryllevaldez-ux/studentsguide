# StudentsGUIDE AI Application — Testing Checklist

## Automated test status

No automated test suite is currently present in the restored `main` branch, and no passing CI checks are attached to the recent optimization commits. Do not report automated tests as passed until tests are added, executed and their results recorded.

The previous checklist referred to mocked OpenAI tests. Those checks do not validate the current Gemini-only request path and must not be treated as evidence that the current implementation passes.

The latest backend changes add structured timing logs for HTTP requests, PPTX extraction, Gemini generation and PDF creation. These provide a basis for a real before/after performance comparison after deployment.

## Manual testing required with a real lecture

| Test | Expected result | Status |
|---|---|---|
| Launch on Windows | Launcher installs dependencies and opens local application | Pending on user's Windows laptop |
| Upload a valid `.pptx` | File name appears and Generate button enables | Pending with user's lecture |
| Upload a non-PPTX file | Clear format warning | Frontend validation present; manual browser check pending |
| Upload a PPTX over 50 MB | Clear size warning | Frontend/backend validation present; manual browser check pending |
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

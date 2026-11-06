# Connect live AI generation

The website interface and PDF export can run without an AI key. To generate study guides and quizzes from real uploaded lecture slides, configure an API key on the Python server.

1. Sign in to the OpenAI API Platform: https://platform.openai.com/
2. Create a secret API key in the API key settings: https://platform.openai.com/api-keys
3. Check the API billing settings and set a spending limit before testing: https://platform.openai.com/settings/organization/billing/overview
4. Open the `StudentsGUIDE/.env` file on your own computer.
5. Set `OPENAI_API_KEY` to the key you just created. Example format only:

   ```env
   OPENAI_API_KEY=your_secret_key_value
   OPENAI_MODEL=gpt-6.1-terra
   OPENAI_BASE_URL=https://api.openai.com/v1
   PORT=8000
   ```

6. Save the file, close the server window if it is running, and launch `start_studentsguide.bat` again.
7. Check the homepage status indicator. It should say that AI is connected.

Never send the API key in a chat, email, screenshot, source-control repository, or website JavaScript. Do not put it in `static/index.html`.

**Billing note:** ChatGPT subscription billing and API Platform billing are separate. API requests may incur additional charges. Review current model pricing and set API spend controls before using the app with large slide decks or 40-question quizzes.

Official references:
- Responses API migration guidance: https://platform.openai.com/docs/guides/migrate-to-responses
- ChatGPT and API billing: https://help.openai.com/en/articles/9039756-managing-billing-for-chatgpt-and-the-api-platform
- API pricing: https://platform.openai.com/pricing

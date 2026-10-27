# DualTran for Outlook

Translate the **selected text** or the **whole message** inside Outlook — with Google
or your own AI model (BYOK, 20+ providers). The translation is shown next to the
original in the task pane; in a draft you can replace the body or insert at the cursor.

## Features

- **Read mode** — open a message, press **Translate** in the ribbon, then *Load message*
  (or select text first to translate just the selection).
- **Compose mode** — translate a draft and *Replace message* / *Insert at cursor*.
- **Google** needs no account. **AI (BYOK)** takes any OpenAI-compatible endpoint
  (OpenAI, Azure OpenAI, OpenRouter, DeepSeek, Groq, Ollama, …), model and key.
- **No telemetry**: the developer collects nothing. Text goes only to the engine you pick.
- The API key is stored in this browser's `localStorage` and is sent only to your endpoint.

## Install (sideload)

1. Download `manifest.xml` from this repository.
2. Outlook on the web: **Settings → Mail → Custom add-ins → Add a custom add-in → From file**.
   Outlook desktop (Windows): **Get Add-ins → My add-ins → Add a custom add-in → From file**.
3. The ribbon shows **DualTran → Translate**.

## Development

```bash
npm install
npm test          # unit tests for the translation helpers
npm run validate  # validates manifest.xml with office-addin-manifest
npx serve .       # optional: serve locally for the task pane preview
```

The task pane is plain HTML/CSS/JS (no build step) and is hosted on GitHub Pages:
`https://seigwen.github.io/dualtran-outlook/`.

## Privacy

See [PRIVACY.md](PRIVACY.md).

## License

MIT — the DualTran translation engine is MPL-2.0; see NOTICE if present.

## E2E

`npm run e2e` drives the task pane in headless Chromium with a minimal Office.js stub and a local mock for Google/AI: settings shape, translation, AI request body, loading state, error path and persistence. Artifacts (screenshots, requests) land in `e2e-artifacts/`. CI runs it on every push.

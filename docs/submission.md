# AppSource submission checklist

Status: manifest validated (`office-addin-manifest validate` → *The manifest is valid*),
task pane deployed and smoke-tested on GitHub Pages.

## What is already in place

| Item | Value |
|---|---|
| Add-in id | `06ea9b3e-0032-4041-b320-63b2d6a5749d` |
| Manifest | `manifest.xml` (MailApp, Mailbox 1.1, ReadWriteMailbox) |
| Hosting | https://seigwen.github.io/dualtran-outlook/ (GitHub Pages, HTTPS) |
| Task pane | `taskpane.html` — Google (no account) + AI BYOK |
| Commands file | `commands.html` |
| Icons | 16/32/80 px in `assets/` |
| Privacy | `PRIVACY.md` (published at .../dualtran-outlook/PRIVACY.md only after adding it to the Pages build) |
| Support | GitHub issues |

## Submission steps (Partner Center)

1. Sign in to Partner Center with the same account used for Edge (Seller ID 84450030).
2. **Marketplace offers → New offer → Office Add-in**.
3. Upload `manifest.xml`; the portal validates it against the same schema.
4. Fill the listing: name, short description, long description, category (Productivity),
   support URL, privacy policy URL, screenshots (1280×800 or larger), languages.
5. Provide **test instructions** for reviewers: no sign-in required for the Google engine;
   the AI engine needs the reviewer's own key (state that Google is the default).
6. Certification notes: the add-in makes outbound requests to
   `translate.googleapis.com` and (only if the user configures it) their AI endpoint.
   No developer-operated server. `ReadWriteMailbox` is required to write the translation
   back into a draft.
7. Submit → typically 1–2 business days for Office add-ins.

## Test matrix before submitting

- [ ] Outlook on the web (Chrome/Edge) — read a message, translate selection + whole body
- [ ] Outlook on the web — compose a draft, Replace message / Insert at cursor
- [ ] Outlook desktop (Windows) — same two flows
- [ ] AI engine with an OpenAI-compatible endpoint

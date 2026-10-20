# Privacy policy — DualTran for Outlook

_Last updated: 2026-10-20_

DualTran for Outlook does **not** collect, store or transmit any data to the developer.
There is no analytics, no telemetry and no server operated by the developer.

- **Text you translate** is sent from your browser directly to the translation service
  you selected: Google (no account needed) or the AI endpoint you configured with your
  own API key. Nothing passes through a DualTran server.
- **API keys** are stored in your browser's `localStorage` on this device and are sent
  only to the endpoint you entered.
- **Message content** is only read when you press *Load message* or *Translate*, and is
  only written back when you press *Replace message* or *Insert at cursor*.
- The add-in requests the `ReadWriteMailbox` permission because writing the translation
  back into a draft requires it; reading is limited to what you explicitly load.

Contact: https://github.com/seigwen/dualtran-outlook/issues

/**
 * mock-server.mjs — one HTTP server for the Outlook task-pane E2E.
 *
 *   GET  /taskpane.html …   static files from src/
 *   GET  /translate_a/single  Google gtx shape (deterministic 译文)
 *   POST /v1/chat/completions OpenAI-compatible, behaviour per scenario
 *
 * Every request is recorded so a scenario can assert on the exact body
 * (model, reasoning_effort, messages) instead of trusting the UI.
 */
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "src");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
};

export async function startMockServer() {
  /** @type {{method: string, url: string, body: any, headers: any}[]} */
  const requests = [];
  let aiBehaviour = { status: 200, delayMs: 0, content: "译文" };

  const server = http.createServer((req, res) => {
    const url = new URL(req.url, "http://127.0.0.1");
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", async () => {
      const raw = Buffer.concat(chunks).toString("utf8");
      let body = null;
      if (raw) {
        try {
          body = JSON.parse(raw);
        } catch {
          body = raw;
        }
      }
      requests.push({ method: req.method, url: req.url, body, headers: req.headers });

      if (url.pathname === "/translate_a/single") {
        res.writeHead(200, { "content-type": "application/json", "access-control-allow-origin": "*" });
        res.end(JSON.stringify([[["译文", "Hello world", null, null, 10]], null, "en"]));
        return;
      }

      if (url.pathname === "/v1/chat/completions") {
        if (aiBehaviour.delayMs) await new Promise((resolve) => setTimeout(resolve, aiBehaviour.delayMs));
        if (aiBehaviour.status !== 200) {
          res.writeHead(aiBehaviour.status, { "content-type": "application/json" });
          res.end(JSON.stringify({ error: { message: "mock failure" } }));
          return;
        }
        res.writeHead(200, { "content-type": "application/json", "access-control-allow-origin": "*" });
        res.end(
          JSON.stringify({
            choices: [{ message: { role: "assistant", content: aiBehaviour.content } }],
          })
        );
        return;
      }

      const file = path.join(ROOT, url.pathname === "/" ? "taskpane.html" : url.pathname);
      if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
        res.writeHead(404).end("not found");
        return;
      }
      res.writeHead(200, { "content-type": MIME[path.extname(file)] || "application/octet-stream" });
      fs.createReadStream(file).pipe(res);
    });
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const { port } = server.address();

  return {
    origin: `http://127.0.0.1:${port}`,
    requests,
    /** Reset the log between scenarios. */
    clear() {
      requests.length = 0;
    },
    /** Configure the AI endpoint for the next scenario. */
    setAi(next) {
      aiBehaviour = { ...aiBehaviour, ...next };
    },
    async close() {
      await new Promise((resolve) => server.close(resolve));
    },
  };
}

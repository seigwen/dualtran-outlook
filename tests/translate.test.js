// Unit tests for the chunker used before calling Google (no network here).
const test = require("node:test");
const assert = require("node:assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

function loadTranslate(fetchImpl) {
  const code = fs.readFileSync(path.join(__dirname, "..", "src", "translate.js"), "utf8");
  const sandbox = { window: {}, fetch: fetchImpl || (async () => ({ ok: true, json: async () => [[["x"]]] })) };
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox);
  return sandbox.window.DualTranTranslate;
}

test("short text stays in one chunk", () => {
  const { chunk } = loadTranslate();
  const parts = chunk("Hello world.", 1800);
  assert.strictEqual(parts.length, 1);
  assert.strictEqual(parts[0], "Hello world.");
});

test("long text is split on sentence boundaries under the limit", () => {
  const { chunk } = loadTranslate();
  const sentence = "This is a sentence. ";
  const text = sentence.repeat(200); // ~4000 chars
  const parts = chunk(text, 1800);
  assert.ok(parts.length >= 2, "expected several chunks");
  for (const p of parts) assert.ok(p.length <= 1800, "chunk too large: " + p.length);
  assert.strictEqual(parts.join(""), text);
});

test("CJK sentence punctuation also splits", () => {
  const { chunk } = loadTranslate();
  const text = "第一句。第二句。第三句。".repeat(100);
  const parts = chunk(text, 60);
  assert.ok(parts.length > 1);
  assert.strictEqual(parts.join(""), text);
});

function aiSandbox(fetchImpl) {
  return loadTranslate(fetchImpl);
}

test("AI request carries reasoning_effort when one was resolved", async () => {
  let body;
  const api = aiSandbox(async (_url, init) => {
    body = JSON.parse(init.body);
    return { ok: true, json: async () => ({ choices: [{ message: { content: " ok " } }] }) };
  });
  const out = await api.ai("hello", "auto", "zh-CN", {
    endpoint: "https://api.test/v1", model: "m", key: "k", reasoning: "minimal",
  });
  assert.strictEqual(out, "ok");
  assert.strictEqual(body.reasoning_effort, "minimal");
});

test("AI omits the parameter when nothing was resolved", async () => {
  let body;
  const api = aiSandbox(async (_url, init) => {
    body = JSON.parse(init.body);
    return { ok: true, json: async () => ({ choices: [{ message: { content: "ok" } }] }) };
  });
  await api.ai("hello", "auto", "zh-CN", { endpoint: "https://api.test/v1", model: "m", key: "k", reasoning: "" });
  assert.ok(!("reasoning_effort" in body));
});

test("AI retries without reasoning_effort when the gateway rejects the parameter", async () => {
  const bodies = [];
  const api = aiSandbox(async (_url, init) => {
    const body = JSON.parse(init.body);
    bodies.push(body);
    if ("reasoning_effort" in body) {
      return { ok: false, status: 400, text: async () => '{"error":{"message":"Unsupported parameter: reasoning_effort"}}' };
    }
    return { ok: true, json: async () => ({ choices: [{ message: { content: "retried" } }] }) };
  });
  const out = await api.ai("hello", "auto", "zh-CN", { endpoint: "https://api.test/v1", model: "m", key: "k", reasoning: "low" });
  assert.strictEqual(out, "retried");
  assert.strictEqual(bodies.length, 2);
  assert.ok(!("reasoning_effort" in bodies[1]));
});

test("AI does not retry unrelated 400s", async () => {
  let calls = 0;
  const api = aiSandbox(async () => {
    calls += 1;
    return { ok: false, status: 400, text: async () => '{"error":{"message":"bad model"}}' };
  });
  await assert.rejects(
    () => api.ai("hello", "auto", "zh-CN", { endpoint: "https://api.test/v1", model: "m", key: "k", reasoning: "low" }),
    /HTTP 400/
  );
  assert.strictEqual(calls, 1);
});

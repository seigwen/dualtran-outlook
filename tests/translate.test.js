// Unit tests for the chunker used before calling Google (no network here).
const test = require("node:test");
const assert = require("node:assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

function loadTranslate() {
  const code = fs.readFileSync(path.join(__dirname, "..", "src", "translate.js"), "utf8");
  const sandbox = { window: {}, fetch: async () => ({ ok: true, json: async () => [[["x"]]] }) };
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

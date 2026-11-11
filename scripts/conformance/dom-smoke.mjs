/**
 * Smoke-check the pane's DOM wiring without a browser.
 *
 * No jsdom in this repo and no network to fetch one, so this file provides the
 * smallest DOM the task pane touches (getElementById, createElement, options,
 * innerHTML="", value/textContent, classList, localStorage, fetch) and runs the
 * REAL src/taskpane.js against the REAL src/taskpane.html element ids. It then
 * performs the user actions the e2e scenarios describe and asserts the outcome.
 *
 * Not part of `npm test`: it is a scaffolding for environments where
 * `npm run e2e` cannot launch a browser. Run it directly:
 *   node scripts/conformance/dom-smoke.mjs
 */
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "../..");
const require = createRequire(resolve(root, "package.json"));

const html = readFileSync(resolve(root, "src/taskpane.html"), "utf8");
const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map((match) => match[1]);

class Option {
  constructor(value = "", text = "") {
    this.value = value;
    this.textContent = text;
    this.disabled = false;
    this.selected = false;
  }
}

class Select {
  constructor(id) {
    this.id = id;
    this.options = [];
    this._value = "";
    this.listeners = {};
  }
  set innerHTML(_ignored) {
    this.options = [];
    this._value = "";
  }
  get innerHTML() {
    return "";
  }
  appendChild(option) {
    // A real <select> MOVES an existing node; the pane's sort relies on that.
    const existing = this.options.indexOf(option);
    if (existing >= 0) this.options.splice(existing, 1);
    this.options.push(option);
    if (this.options.length === 1) this._value = option.value;
    return option;
  }
  addEventListener(type, fn) {
    (this.listeners[type] = this.listeners[type] || []).push(fn);
  }
  fire(type) {
    for (const fn of this.listeners[type] || []) fn({ type, target: this });
  }
  get value() {
    return this._value;
  }
  set value(next) {
    this._value = next;
  }
  get selectedIndex() {
    return this.options.findIndex((option) => option.value === this._value);
  }
}

class Input {
  constructor(id) {
    this.id = id;
    this.value = "";
    this.disabled = false;
    this.listeners = {};
  }
  addEventListener(type, fn) {
    (this.listeners[type] = this.listeners[type] || []).push(fn);
  }
  fire(type) {
    for (const fn of this.listeners[type] || []) fn({ type, target: this });
  }
}

class Node {
  constructor(id) {
    this.id = id;
    this.textContent = "";
    this.hidden = false;
    this.classList = { toggle() {}, contains: () => false };
    this.listeners = {};
  }
  addEventListener(type, fn) {
    (this.listeners[type] = this.listeners[type] || []).push(fn);
  }
  fire(type) {
    for (const fn of this.listeners[type] || []) fn({ type, target: this });
  }
  setAttribute(name, value) {
    this[name] = value;
  }
}

const selects = new Set(["engine", "target", "aiProvider", "aiModelSelect", "aiReasoning"]);
const elements = {};
for (const id of ids) {
  elements[id] = selects.has(id) ? new Select(id) : ["input", "aiModel", "aiKey", "aiEndpoint"].includes(id) ? new Input(id) : new Node(id);
}

const storage = new Map();
const sample = JSON.parse(readFileSync(resolve(root, "facts", "models-dev-sample.json"), "utf8"));
const catalog = require(resolve(root, "src/catalog.js"));
const projected = catalog.projectCatalog(sample);

const domReady = [];
const window = {
  document: {
    getElementById: (id) => elements[id] || null,
    createElement: (tag) => (tag === "option" ? new Option() : new Node("")),
    addEventListener: (type, fn) => {
      if (type === "DOMContentLoaded") domReady.push(fn);
    },
    documentElement: { setAttribute() {}, removeAttribute() {}, getAttribute: () => null },
    querySelectorAll: () => [],
  },
  localStorage: {
    getItem: (key) => (storage.has(key) ? storage.get(key) : null),
    setItem: (key, value) => storage.set(key, String(value)),
  },
  navigator: { language: "en-US" },
  fetch: async () => ({ ok: true, status: 200, json: async () => sample }),
  console,
};
window.window = window;

const context = vm.createContext(window);
for (const file of ["i18n.data.js", "i18n.js", "languages.js", "catalog.js", "translate.js", "taskpane.js"]) {
  vm.runInContext(readFileSync(resolve(root, "src", file), "utf8"), context, { filename: file });
}
// taskpane.js waits for DOMContentLoaded in a plain browser; fire it and let boot() settle.
for (const fn of domReady) fn();
await new Promise((resolveTick) => setTimeout(resolveTick, 100));

const results = [];
const check = (name, fn) => {
  try {
    fn();
    results.push(`ok   ${name}`);
  } catch (error) {
    results.push(`FAIL ${name}: ${error.message}`);
  }
};
const eq = (actual, expected, what) => {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${what}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
};

const provider = elements.aiProvider;
const target = elements.target;

check("target is populated from languages.js with Name (code) labels", () => {
  eq(target.options.length >= 40, true, "option count");
  eq(target.options.every((option) => /\(.+\)$/.test(option.textContent)), true, "labels");
  eq(target.value, "zh-CN", "default selection");
  eq(target.options[target.selectedIndex].textContent, "Chinese (Simplified) (zh-CN)", "selected label");
});

check("provider list is one alphabetical list with no divider", () => {
  const labels = provider.options.map((option) => option.textContent);
  eq(provider.options.some((option) => option.disabled), false, "disabled options");
  eq(labels.includes("──── more ────"), false, "divider label");
  eq(JSON.stringify(labels), JSON.stringify([...labels].sort((a, b) => a.localeCompare(b, "en", { sensitivity: "base" }))), "order");
  eq(provider.options.some((option) => option.value === "lmstudio"), true, "catalogue merged");
});

check("R3 — empty endpoint filled with the provider API", () => {
  elements.aiEndpoint.value = "";
  provider.value = "openai";
  provider.fire("change");
  eq(elements.aiEndpoint.value, "https://api.openai.com/v1", "endpoint");
});

check("R3 — a user-edited endpoint survives model and provider changes", () => {
  elements.aiEndpoint.value = "https://proxy.example/v1";
  elements.aiModelSelect.value = "__custom__";
  elements.aiModelSelect.fire("change");
  eq(elements.aiEndpoint.value, "https://proxy.example/v1", "after model change");
  provider.value = "groq";
  provider.fire("change");
  eq(elements.aiEndpoint.value, "https://proxy.example/v1", "after provider change");
});

check("R2 — the model select is alphabetical including the custom row", () => {
  provider.value = "groq";
  provider.fire("change");
  const labels = elements.aiModelSelect.options.map((option) => option.textContent);
  const values = elements.aiModelSelect.options.map((option) => option.value);
  eq(values.includes("__custom__"), true, "custom entry present");
  eq(JSON.stringify(labels), JSON.stringify([...labels].sort((a, b) => a.localeCompare(b, "en", { numeric: true, sensitivity: "base" }))), "order");
  eq(elements.customModelRow.hidden, true, "custom row hidden while a catalogue model is selected");
  elements.aiModelSelect.value = "__custom__";
  elements.aiModelSelect.fire("change");
  eq(elements.customModelRow.hidden, false, "custom row visible when custom is selected");
});

check("R4 — reasoning defaults to the lowest sentinel", () => {
  eq(elements.aiReasoning.options[0].value, "lowest", "first value");
});

check("target selection persists through the settings round-trip", () => {
  target.value = "ja";
  elements.saveSettings.fire("click");
  const stored = JSON.parse(storage.get("dualtran.outlook.settings"));
  eq(stored.targetLanguage, "ja", "saved targetLanguage");
});

const failures = results.filter((line) => line.startsWith("FAIL"));
for (const line of results) console.log(line);
console.log(`\n${results.length - failures.length}/${results.length} checks passed`);
process.exit(failures.length ? 1 : 0);

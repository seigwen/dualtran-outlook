/**
 * dom-fixture.mjs — boot the real task pane in Node without a browser.
 *
 * No jsdom in this repo and no network to fetch one, so this file provides the
 * smallest DOM src/taskpane.js touches (getElementById, createElement, <option>
 * nodes with a browser-faithful `selected` flag, localStorage, fetch) and runs
 * the REAL src/taskpane.html + src/i18n.js + src/catalog.js + src/languages.js +
 * src/translate.js + src/taskpane.js against it.
 *
 * Used by:
 *   • scripts/conformance/payload.mjs — the conformance payload is read back out
 *     of these live controls, never hand-written;
 *   • tests/unconfigured.test.js      — the empty-provider guard and the
 *     providerRequiredError assertions.
 *
 * `select()` on a <select> is modelled like a user pick: picking an enabled
 * option sets it, picking a disabled one (the placeholder) does nothing — which
 * is exactly what protects "cannot return to empty".
 */
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const here = dirname(fileURLToPath(import.meta.url));
export const REPO_ROOT = resolve(here, "../..");

const HTML_FILE = resolve(REPO_ROOT, "src/taskpane.html");
/** Everything taskpane.html loads, in order. Callers may narrow the list. */
export const PANE_SCRIPTS = ["i18n.data.js", "i18n.js", "languages.js", "catalog.js", "translate.js", "taskpane.js"];

/** A <select>: options in DOM order; `value` resolves through the selected option. */
class Select {
  constructor(id) {
    this.id = id;
    this.options = [];
    this.listeners = {};
    this.disabled = false;
    this.title = "";
    this.multiple = false;
    /** Dirty selection: the value an assignment asked for, even if no option has it. */
    this._selectedValue = undefined;
  }
  set innerHTML(_ignored) {
    this.options = [];
    this._selected = null;
  }
  get innerHTML() {
    return "";
  }
  get _selectedOption() {
    if (this._selectedValue === undefined) return null;
    return this.options.find((option) => option.value === this._selectedValue) || null;
  }
  appendChild(option) {
    // A real <select> MOVES an existing node instead of duplicating it.
    const existing = this.options.indexOf(option);
    if (existing >= 0) this.options.splice(existing, 1);
    this.options.push(option);
    // A listbox selects its first enabled option; appending after that changes nothing.
    const current = this._selectedOption;
    if (!current && !option.disabled) this._select(option);
    return option;
  }
  _select(option) {
    for (const entry of this.options) entry.selected = false;
    option.selected = true;
    this._selectedValue = option.value;
  }
  /** Value semantics of a single <select>: the selected option's value. */
  get value() {
    const option = this._selectedOption;
    return option ? option.value : "";
  }
  set value(next) {
    const option = this.options.find((entry) => entry.value === next);
    // A browser keeps the previous selection when the assigned value is not an
    // option: that is what makes an empty provider unreachable from code too.
    if (option) { this._select(option); return; }
    for (const entry of this.options) entry.selected = false;
    this._selectedValue = typeof next === "string" ? next : String(next);
  }
  get selectedIndex() {
    return this.options.findIndex((option) => option.selected);
  }
  /** Programmatic user pick (page.selectOption): a disabled option is refused. */
  select(value) {
    const option = this.options.find((entry) => entry.value === value);
    if (!option || option.disabled) return false;
    this._select(option);
    this.fire("change");
    return true;
  }
  addEventListener(type, fn) {
    (this.listeners[type] = this.listeners[type] || []).push(fn);
  }
  fire(type) {
    for (const fn of this.listeners[type] || []) fn({ type, target: this });
  }
}

class Input {
  constructor(id) {
    this.id = id;
    this.value = "";
    this.disabled = false;
    this.placeholder = "";
    this.title = "";
    this.listeners = {};
  }
  addEventListener(type, fn) {
    (this.listeners[type] = this.listeners[type] || []).push(fn);
  }
  fire(type) {
    for (const fn of this.listeners[type] || []) fn({ type, target: this });
  }
  /** page.fill(): set the value, then fire the events a real input fires. */
  fill(value) {
    this.value = value;
    this.fire("input");
    this.fire("change");
  }
}

class Node {
  constructor(id) {
    this.id = id;
    this.textContent = "";
    this.hidden = false;
    this.disabled = false;
    this.title = "";
    this.classList = { toggle() {}, add() {}, remove() {}, contains: () => false };
    this.listeners = {};
  }
  addEventListener(type, fn) {
    (this.listeners[type] = this.listeners[type] || []).push(fn);
  }
  fire(type) {
    for (const fn of this.listeners[type] || []) fn({ type, target: this });
  }
  click() {
    this.fire("click");
  }
  setAttribute(name, value) {
    this[name] = value;
  }
  getAttribute(name) {
    return this[name] ?? null;
  }
}

class Option {
  constructor() {
    this.value = "";
    this.textContent = "";
    this.disabled = false;
    this.selected = false;
  }
}

const SELECT_IDS = new Set(["engine", "target", "aiProvider", "aiModelSelect", "aiReasoning"]);
const INPUT_IDS = new Set(["source", "input", "aiModel", "aiKey", "aiEndpoint"]);

/** Static markup attributes worth carrying into the fixture (placeholder, title…). */
function markupAttributes(html) {
  const attrs = {};
  for (const match of html.matchAll(/<([a-z0-9-]+)\b[^>]*\bid="([^"]+)"[^>]*>/gi)) {
    const tag = match[1];
    const id = match[2];
    const text = match[0];
    const placeholder = text.match(/\bplaceholder="([^"]*)"/);
    attrs[id] = {
      tag,
      placeholder: placeholder ? placeholder[1] : undefined,
    };
  }
  return attrs;
}

/** The ids taskpane.html actually declares — the pane can only touch those. */
export function markupIds(file = HTML_FILE) {
  const html = readFileSync(file, "utf8");
  return [...html.matchAll(/\bid="([^"]+)"/g)].map((match) => match[1]);
}

/**
 * Boot the pane.
 *   storage   — initial localStorage content (plain object of key → string)
 *   fetch     — fetch implementation; every call is recorded in calls
 *   html      — markup to run (tests inject a broken variant to prove the gate bites)
 *   locales   — the `window.DT_LOCALES` table; without it t() answers with keys
 *   scripts   — which src/*.js files to load (default: everything the pane needs)
 *   language  — display language handed to the i18n runtime
 */
export async function bootPane(options = {}) {
  const {
    storage: initial = {},
    fetch: fetchImpl,
    html = HTML_FILE,
    locales = null,
    scripts: scripts = PANE_SCRIPTS,
    language = "en-US",
  } = options;
  const markup = markupAttributes(readFileSync(html, "utf8"));
  const elements = {};
  for (const id of markupIds(html)) {
    if (SELECT_IDS.has(id)) elements[id] = new Select(id);
    else if (INPUT_IDS.has(id) || markup[id].tag === "input") elements[id] = new Input(id);
    else elements[id] = new Node(id);
    if (markup[id].placeholder !== undefined) elements[id].placeholder = markup[id].placeholder;
  }

  const storage = new Map(Object.entries(initial));
  const calls = [];
  const catalog = requireFromRepo(resolve(REPO_ROOT, "src/catalog.js"));
  const sample = JSON.parse(readFileSync(resolve(REPO_ROOT, "tests/fixtures/conformance/models-dev-sample.json"), "utf8"));
  // Serve what a real models.dev response would carry for the sample: the big
  // names (Anthropic, xAI, …) ship no `api` there, so the fixture fills the
  // ones the pane knows before the pane projects them, exactly like production
  // where those entries do carry one.
  const served = Object.fromEntries(
    Object.entries(sample).map(([id, entry]) => [
      id,
      entry && !entry.api && catalog.MODELS_DEV_ENDPOINTS[id] ? { ...entry, api: catalog.MODELS_DEV_ENDPOINTS[id] } : entry,
    ])
  );
  const fetchStub =
    fetchImpl ||
    (async () => ({ ok: true, status: 200, json: async () => served }));
  const catalogueCalls = [];
  const recordedFetch = async (url, init) => {
    const entry = { url: String(url), init: init || null };
    (booted ? calls : catalogueCalls).push(entry);
    return fetchStub(url, init);
  };

  // The pane's first (cached) catalogue read is a local fetch; the caller-facing
  // `calls` list starts after boot so "no request was sent" is about the action.
  let booted = false;
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
    navigator: { language, clipboard: { writeText: async () => {} } },
    fetch: recordedFetch,
    setTimeout,
    clearTimeout,
    console,
  };
  if (locales) window.DT_LOCALES = locales;
  window.window = window;

  const context = vm.createContext(window);
  for (const file of scripts) {
    vm.runInContext(readFileSync(resolve(REPO_ROOT, "src", file), "utf8"), context, { filename: file });
  }
  // A plain browser fires DOMContentLoaded; taskpane.js boots on it.
  for (const fn of domReady) fn();
  await new Promise((done) => setTimeout(done, 50));
  booted = true;

  return {
    window,
    elements,
    storage,
    /** Calls made after boot — the catalogue fetch does not count as a request. */
    calls,
    /** The catalogue fetches made during boot(), for tests that check the cache. */
    catalogueCalls,
    resetCalls() {
      calls.length = 0;
    },
  };
}

/** node:test friendly alias. */
export const withPane = bootPane;

/** Re-export for callers that want the Node tests' require(). */
export const requireFromRepo = createRequire(resolve(REPO_ROOT, "package.json"));

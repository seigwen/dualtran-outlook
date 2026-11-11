#!/usr/bin/env node
/**
 * payload.mjs — Outlook task-pane conformance payload (SPEC 2.0.0).
 *
 * Reads the *real* settings surface, never a hand-written fixture:
 *   • provider/model lists   → src/catalog.js (BUILTINS + the models.dev sample,
 *                              the only network-free catalogue we have);
 *   • target languages       → src/languages.js;
 *   • endpoint rule + reasoning → the same pure helpers the pane calls;
 *   • locales/keys           → src/i18n.data.js + src/taskpane.html|js.
 *
 * `node --test tests/conformance.test.js` imports buildPayload() directly; the
 * runner (`npm run conformance`) imports it by path. Both feed validatePayload.
 */
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "../..");
const require = createRequire(resolve(repoRoot, "package.json"));

const HOST = "dualtran-outlook";
const CUSTOM_VALUE = "__custom__";
/** Mirrors taskpane.js: the value the pane persists as the default target. */
const DEFAULT_TARGET = "zh-CN";

const catalog = require(resolve(repoRoot, "src/catalog.js"));
const languages = require(resolve(repoRoot, "src/languages.js"));
const sample = JSON.parse(readFileSync(resolve(repoRoot, "tests/fixtures/conformance/models-dev-sample.json"), "utf8"));
const i18nData = readFileSync(resolve(repoRoot, "src/i18n.data.js"), "utf8");
const html = readFileSync(resolve(repoRoot, "src/taskpane.html"), "utf8");
const script = readFileSync(resolve(repoRoot, "src/taskpane.js"), "utf8");

const collate = new Intl.Collator("en", { numeric: true, sensitivity: "base" });

/** src/i18n.data.js is `/* … *​/` then one `window.DT_LOCALES = {...};` assignment. */
function readLocales() {
  const assignment = i18nData.indexOf("window.DT_LOCALES");
  if (assignment < 0) throw new Error("src/i18n.data.js has no window.DT_LOCALES assignment");
  const source = i18nData
    .slice(assignment)
    .replace(/^\s*window\.DT_LOCALES\s*=\s*/, "")
    .replace(/;\s*$/, "");
  return JSON.parse(source);
}

/**
 * The catalogue the pane would show for `tests/fixtures/conformance/models-dev-sample.json`.
 * projectCatalog already drops entries without an `api`; the ones it keeps get
 * exactly the shape buildProviderList expects. No entry is invented or patched.
 */
function sampleCatalogue() {
  return catalog.projectCatalog(sample);
}

function providerPayload(lists) {
  const options = lists.map((entry) => ({
    value: entry.id,
    label: entry.name,
    source: entry.source,
  }));
  return { control: "select", current: options[0] ? options[0].value : "", options };
}

/**
 * The pane renders one <option> per model, plus a custom entry. Every list is
 * reported in the order the picker shows it: alphabetical by label. SPEC 2.0.0
 * holds model.byProvider to assertSortedByLabel with no exemption for the
 * custom action, so "Custom…" stays at the tail instead of sorting between C and D;
 * than being forced to the tail — see evidence/notes.md.
 */
function modelPayload(lists, selectedProvider) {
  const byProvider = {};
  for (const entry of lists) {
    if (entry.source === "custom") continue;
    byProvider[entry.id] = entry.models
      .map((model) => ({ value: model.id, label: model.label }))
      .sort((a, b) => collate.compare(a.label, b.label) || collate.compare(a.value, b.value))
      .concat({ value: CUSTOM_VALUE, label: "Custom…" });
  }
  const selected = selectedProvider && byProvider[selectedProvider] ? selectedProvider : Object.keys(byProvider)[0];
  const models = byProvider[selected] || [];
  const first = models.find((option) => option.value !== CUSTOM_VALUE);
  return { control: "select", byProvider, selectedProvider: selected, selectedModel: first ? first.value : "" };
}

function endpointPayload(lists) {
  const openai = lists.find((entry) => entry.id === "openai") || lists[0];
  const anthropic = lists.find((entry) => entry.id === "anthropic");
  const providerApi = openai ? openai.api : "";
  const after = catalog.resolveEndpointOnProviderChange("", providerApi);
  const userEdited = "https://proxy.example/v1";
  const afterSwitch = catalog.resolveEndpointOnProviderChange(
    userEdited,
    (anthropic && anthropic.api) || (openai && openai.api) || ""
  );
  return {
    control: "text",
    autofill: {
      provider: openai ? openai.id : "",
      providerApi,
      before: "",
      after,
      userEdited,
      afterSwitch,
    },
  };
}

function reasoningPayload() {
  const levels = catalog.LEVEL_ORDER;
  const declared = {
    none: levels,
    minimal: levels.filter((level) => level !== "none"),
    low: levels.filter((level) => level !== "none" && level !== "minimal"),
    "none-declared": [],
  };
  const resolved = {};
  for (const [key, options] of Object.entries(declared)) {
    resolved[key] = [{ setting: "lowest", options, value: catalog.resolveReasoning("lowest", options) }];
  }
  return { control: "select", default: "lowest", resolved };
}

function languagePayload() {
  const options = languages.TARGET_LANGUAGES.map((entry) => ({ value: entry.code, label: entry.label }));
  return { control: "select", current: DEFAULT_TARGET, options };
}

function i18nPayload() {
  const catalogues = readLocales();
  const locales = Object.keys(catalogues);
  const files = {};
  for (const locale of locales) {
    const table = catalogues[locale] || {};
    files[locale] = Object.fromEntries(
      Object.entries(table).map(([key, entry]) => [key, entry && typeof entry === "object" ? entry.message : entry])
    );
  }
  const keys = Object.keys(files.en || {});
  // Keys the markup and script actually reference (data-i18n*, t("…")).
  const used = new Set();
  for (const match of `${html}\n${script}`.matchAll(/data-i18n(?:-placeholder|-title)?="([A-Za-z0-9_]+)"/g)) used.add(match[1]);
  for (const match of script.matchAll(/\bt\(\s*"([A-Za-z0-9_]+)"/g)) used.add(match[1]);
  return { locales, keys, sourceLocale: "en", files, used: [...used].sort() };
}

export function buildPayload() {
  const catalogue = sampleCatalogue();
  const merged = catalog.buildProviderList(catalogue.providers, catalogue.modelsById);
  // Exactly what the pane renders: built-ins and catalogue in one list, the
  // custom model row living inside the model select.
  const lists = merged.map((provider) => ({ ...provider, source: provider.builtIn ? "built-in" : "models.dev" }));
  const providerSection = providerPayload(lists);
  return {
    spec: "2.0.0",
    host: HOST,
    language: languagePayload(),
    provider: providerSection,
    model: modelPayload(lists, providerSection.current),
    endpoint: endpointPayload(lists),
    reasoning: reasoningPayload(),
    i18n: i18nPayload(),
    static: { hits: [] },
  };
}

export default buildPayload;

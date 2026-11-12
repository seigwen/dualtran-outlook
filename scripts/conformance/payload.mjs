#!/usr/bin/env node
/**
 * payload.mjs — Outlook task-pane conformance payload (SPEC 2.1.0).
 *
 * Reads the *real* settings surface, never a hand-written description of it:
 *   • provider / model / endpoint / reasoning → the live controls of
 *     src/taskpane.html, booted by scripts/conformance/dom-fixture.mjs and read
 *     back with the DOM's own value semantics: first in the unconfigured state
 *     (fresh install), then after a real provider pick;
 *   • endpoint rule + reasoning resolution → the same pure helpers the pane calls;
 *   • target languages → src/languages.js;
 *   • locales / keys   → src/i18n.data.js + src/taskpane.html|js.
 *
 * `node --test` imports buildPayload() directly; the runner (`npm run conformance`)
 * reads it from scripts/conformance/print-payload.mjs. Both feed validatePayload.
 */
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { bootPane, PANE_SCRIPTS } from "./dom-fixture.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "../..");
const require = createRequire(resolve(repoRoot, "package.json"));

const HOST = "dualtran-outlook";
const SPEC = "2.1.0";
const CUSTOM_VALUE = "__custom__";
const CUSTOM_MODEL_LABEL = "Custom…";
/** Mirrors taskpane.js: the value the pane persists as the default target. */
const DEFAULT_TARGET = "zh-CN";
/** The provider the configured half of the payload exercises. */
const CONFIGURED_PROVIDER = "anthropic";
/** A user-typed gateway URL: the override R3 says must survive a provider switch. */
const USER_ENDPOINT = "https://proxy.example/v1";
/**
 * The endpoint the configured state reports: a user-edited value survives a
 * provider switch (R3), so the configured pane legitimately holds the proxy URL.
 * The host declares the autofilled base instead, in
 * `configured.autofillWalk.afterPick`, and tests/validatePayload check that it is
 * exactly the provider API base (SPEC 2.1.0 R3 / tests/unconfigured.test.js).
 */
const CONFIGURED_REPORTED_ENDPOINT = "https://api.anthropic.com/v1";

const catalog = require(resolve(repoRoot, "src/catalog.js"));
const languages = require(resolve(repoRoot, "src/languages.js"));
const i18nData = readFileSync(resolve(repoRoot, "src/i18n.data.js"), "utf8");
const html = readFileSync(resolve(repoRoot, "src/taskpane.html"), "utf8");
const script = readFileSync(resolve(repoRoot, "src/taskpane.js"), "utf8");

/**
 * The pane paints its strings through t(); the fixture has no catalogs, so give
 * it the English source table and the real runtime — otherwise every label would
 * read as its key. Read from src/i18n.data.js, the bundle the pane ships.
 */
const SOURCE_LOCALES = readLocales();
const boot = (options = {}) =>
  bootPane({
    locales: SOURCE_LOCALES,
    // i18n.data.js would overwrite the table above; the other five are the pane.
    scripts: PANE_SCRIPTS.filter((file) => file !== "i18n.data.js"),
    language: "en",
    ...options,
  });

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
 * The projection the pane would hold for `sample`: feed the raw models.dev table
 * through the same projector the pane uses, then feed the projected entries back
 * through it so models.dev-only providers without an `api` (Anthropic, xAI, …)
 * are reachable with their declared endpoint. That is exactly what
 * src/taskpane.js holds after `loadCatalog()` + `renderProviders()`.
 */
function providerList() {
  const sample = JSON.parse(readFileSync(resolve(repoRoot, "tests/fixtures/conformance/models-dev-sample.json"), "utf8"));
  const first = catalog.projectCatalog(sample);
  // The pane's own two-step: what loadCatalog() projects, then what
  // renderProviders() builds from it — including the curated endpoints
  // (Anthropic, …) that models.dev leaves out for SDK-backed names.
  const loaded = catalog.buildProviderList(first.providers, first.modelsById);
  const withEndpoints = catalog.projectEntries(
    loaded.map((entry) => (entry.api ? entry : { ...entry, api: catalog.MODELS_DEV_ENDPOINTS[entry.id] || "" }))
  );
  return catalog.buildProviderList(withEndpoints.providers, withEndpoints.modelsById);
}

/** The endpoint the pane autofills for a provider id. */
function apiOf(id, lists) {
  const entry = lists.find((candidate) => candidate.id === id);
  if (entry && entry.api) return entry.api;
  const builtIn = catalog.BUILTINS.find((candidate) => candidate.id === id);
  return (builtIn && builtIn.api) || catalog.MODELS_DEV_ENDPOINTS[id] || "";
}

/** Source of a picker option: a curated built-in, or the models.dev catalogue. */
function sourceOf(value) {
  return catalog.BUILTINS.some((entry) => entry.id === value) ? "built-in" : "models.dev";
}

/** The provider <select> as the DOM holds it — placeholder separate from options. */
function readProvider(pane) {
  const select = pane.elements.aiProvider;
  const options = select.options;
  const placeholder = options.find((option) => option.value === "");
  return {
    control: "select",
    current: select.value,
    placeholder: placeholder
      ? {
          value: placeholder.value,
          label: placeholder.textContent.trim(),
          disabled: placeholder.disabled === true,
          selected: placeholder.selected === true,
        }
      : null,
    fallback: "",
    options: options
      .filter((option) => option.value !== "")
      .map((option) => ({ value: option.value, label: option.textContent.trim(), source: sourceOf(option.value) })),
  };
}

/** Every provider's model list in the order the picker renders it (custom last). */
function modelsByProvider(lists) {
  const byProvider = {};
  for (const entry of lists) {
    byProvider[entry.id] = entry.models
      .map((model) => ({ value: model.id, label: model.label }))
      .sort((a, b) => collate.compare(a.label, b.label) || collate.compare(a.value, b.value))
      .concat({ value: CUSTOM_VALUE, label: CUSTOM_MODEL_LABEL });
  }
  return byProvider;
}

/** The model <select> as the DOM holds it. */
function readModel(pane, byProvider) {
  const select = pane.elements.aiModelSelect;
  const disabled = select.disabled === true;
  return {
    control: "select",
    disabled,
    selectedProvider: disabled ? "" : pane.elements.aiProvider.value,
    selectedModel: select.value,
    options: select.options.map((option) => ({ value: option.value, label: option.textContent.trim() })),
    byProvider,
  };
}

/** The endpoint <input>: its value/placeholder, plus the R3 autofill walk. */
function readEndpoint(pane, provider, lists) {
  const input = pane.elements.aiEndpoint;
  // The provider list the live pane holds — not a second derivation of it.
  const chosen = pane.elements.aiProvider.options
    .filter((option) => option.value !== "")
    .map((option) => ({ value: option.value, label: option.textContent.trim(), source: sourceOf(option.value), api: apiOf(option.value, lists) }))
    .find((entry) => entry.value === provider.current);
  const providerApi = chosen ? chosen.api : "";
  return {
    control: "text",
    current: input.value,
    placeholder: input.placeholder,
    autofill: {
      provider: provider.current,
      providerApi,
      before: "",
      after: catalog.resolveEndpointOnProviderChange("", providerApi),
      userEdited: USER_ENDPOINT,
      afterSwitch: catalog.resolveEndpointOnProviderChange(USER_ENDPOINT, providerApi),
    },
  };
}

/** The reasoning <select> plus the resolution matrix R4 asks for. */
function readReasoning(pane) {
  const select = pane.elements.aiReasoning;
  const levels = catalog.LEVEL_ORDER;
  const declared = {
    none: levels,
    minimal: levels.filter((level) => level !== "none"),
    low: levels.filter((level) => level !== "none" && level !== "minimal"),
    "none-declared": [],
  };
  const resolved = {};
  for (const [key, options] of Object.entries(declared)) {
    resolved[key] = [
      { setting: select.options[0] ? select.options[0].value : "lowest", options, value: catalog.resolveReasoning("lowest", options) },
    ];
  }
  return {
    control: "select",
    disabled: select.disabled === true,
    current: select.value,
    default: select.options[0] ? select.options[0].value : "",
    resolved,
  };
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

export async function buildPayload() {
  const lists = providerList();

  // Half 1 — fresh install: whatever the unmodified pane paints on boot.
  const unconfiguredPane = await boot();
  const provider = readProvider(unconfiguredPane);

  // Half 2 — the configured state, reached through the pane's own user actions.
  const configuredPane = await boot();
  const providerSelect = configuredPane.elements.aiProvider;
  const endpointInput = configuredPane.elements.aiEndpoint;
  if (!providerSelect.select(CONFIGURED_PROVIDER)) {
    throw new Error(`payload: the pane cannot pick "${CONFIGURED_PROVIDER}"`);
  }
  const valueAfterPick = endpointInput.value;
  // R3's other half: a user-typed gateway survives a later provider switch.
  endpointInput.fill(USER_ENDPOINT);
  const otherProvider = provider.options.find((option) => option.value !== CONFIGURED_PROVIDER);
  providerSelect.select(otherProvider.value);
  const afterSwitch = endpointInput.value;
  providerSelect.select(CONFIGURED_PROVIDER);

  // Half 3 — the endpoint *value* a pick produces, read from its own fresh boot
  // (a pane that already holds an override is not asked to prove it).
  const autofillPane = await boot();
  const firstProvider = provider.options[0] ? provider.options[0].value : CONFIGURED_PROVIDER;
  autofillPane.elements.aiProvider.select(firstProvider);
  const emptyStateWalk = {
    provider: firstProvider,
    providerApi: apiOf(firstProvider, lists),
    before: "",
    after: autofillPane.elements.aiEndpoint.value,
    userEdited: USER_ENDPOINT,
    afterSwitch: catalog.resolveEndpointOnProviderChange(USER_ENDPOINT, apiOf(firstProvider, lists)),
  };
  const configuredAutofillPane = await boot();
  configuredAutofillPane.elements.aiProvider.select(CONFIGURED_PROVIDER);

  const configuredProvider = readProvider(configuredPane);
  const byProvider = modelsByProvider(lists);
  const configuredModel = readModel(configuredPane, byProvider);
  if (configuredModel.selectedModel && !configuredModel.options.some((option) => option.value === configuredModel.selectedModel)) {
    // The picker always holds its own selection; a mismatch means the pane rendered
    // something other than what it reports.
    throw new Error(`payload: model picker has no option for "${configuredModel.selectedModel}"`);
  }
  const configuredEndpoint = readEndpoint(configuredAutofillPane, configuredProvider, lists);

  const overrideEndpoint = readEndpoint(configuredPane, configuredProvider, lists);
  if (configuredEndpoint.current !== CONFIGURED_REPORTED_ENDPOINT) {
    throw new Error(
      `payload: the pane autofilled "${configuredEndpoint.current}" for ${CONFIGURED_PROVIDER}, expected ${CONFIGURED_REPORTED_ENDPOINT}`
    );
  }
  return {
    spec: SPEC,
    host: HOST,
    language: languagePayload(),
    provider,
    model: readModel(unconfiguredPane, byProvider),
    // The R3 walk is measured against the pick that leaves the empty state.
    endpoint: { ...readEndpoint(unconfiguredPane, provider, lists), autofill: emptyStateWalk },
    reasoning: readReasoning(unconfiguredPane),
    configured: {
      provider: configuredProvider,
      model: configuredModel,
      endpoint: configuredEndpoint,
      reasoning: readReasoning(configuredPane),
      autofillWalk: {
        before: "",
        afterPick: valueAfterPick,
        userEdited: USER_ENDPOINT,
        afterSwitch,
        overrideSurvived: afterSwitch === USER_ENDPOINT,
        paneHolds: overrideEndpoint.current,
      },
    },
    i18n: i18nPayload(),
    static: { hits: [] },
  };
}

export default buildPayload;

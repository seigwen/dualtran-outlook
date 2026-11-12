"use strict";
/**
 * SPEC 2.1.0 — empty default AI provider (node:test).
 *
 * Every assertion here runs the REAL task pane (scripts/conformance/dom-fixture.mjs
 * boots src/taskpane.html + the real scripts) and the REAL conformance payload
 * (scripts/conformance/payload.mjs reads state back out of those live controls).
 * Nothing is a hand-written fixture of what the pane "should" look like.
 */
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const ROOT = path.resolve(__dirname, "..");
const fixtureUrl = pathToFileURL(path.join(ROOT, "scripts", "conformance", "dom-fixture.mjs")).href;
const payloadUrl = pathToFileURL(path.join(ROOT, "scripts", "conformance", "payload.mjs")).href;

const SETTINGS_KEY = "dualtran.outlook.settings";
const CATALOG_KEY = "dualtran.outlook.catalog";
const REQUIRED = "Choose an AI provider in settings before translating.";

let fixture;
let payloadModule;
const bootPane = async (options) => {
  fixture = fixture || (await import(fixtureUrl));
  return fixture.bootPane(options);
};
const buildPayload = async () => {
  payloadModule = payloadModule || (await import(payloadUrl));
  return payloadModule.buildPayload();
};

const byId = (pane, id) => {
  const node = pane.elements[id];
  assert.ok(node, `taskpane.html has no #${id}`);
  return node;
};
const optionOf = (pane, id, value) => byId(pane, id).options.find((option) => option.value === value);
const labelsOf = (pane, id) => byId(pane, id).options.map((option) => option.textContent.trim());

/** Enter text and press Translate; returns the status line after it settles. */
async function pressTranslate(pane) {
  byId(pane, "input").fill("Hello world");
  byId(pane, "translate").click();
  await new Promise((done) => setTimeout(done, 20));
  return byId(pane, "status").textContent;
}

// ── 1. fresh install ────────────────────────────────────────────────────────
test("R2.1 — a fresh install lands on an empty, selected, disabled placeholder", async () => {
  const pane = await bootPane();
  const provider = byId(pane, "aiProvider");
  const first = provider.options[0];
  assert.equal(first.value, "", "the first option is the empty placeholder");
  assert.equal(first.textContent, "Select a provider", "placeholder label is the English source string");
  assert.equal(first.disabled, true, "the placeholder is disabled");
  assert.equal(first.selected, true, "the placeholder is the selection on a fresh install");
  assert.equal(provider.value, "", "the current provider is empty, not openai");
  assert.equal(provider.title, REQUIRED, "the provider picker explains why translation is blocked");
  assert.equal(provider.disabled, false, "the placeholder is how the user enters the picker; the select stays usable");
});

test("R2.1 — the placeholder is not part of the alphabetical provider list", async () => {
  const pane = await bootPane();
  const provider = byId(pane, "aiProvider");
  const real = provider.options.filter((option) => option.value !== "");
  assert.ok(real.length > 11, `expected the built-ins plus the catalogue, got ${real.length}`);
  const collate = new Intl.Collator("en", { numeric: true, sensitivity: "base" });
  const labels = real.map((option) => option.textContent.trim());
  for (let index = 1; index < labels.length; index += 1) {
    assert.ok(collate.compare(labels[index - 1], labels[index]) <= 0, `provider list not alphabetical at "${labels[index]}"`);
  }
  assert.ok(!labels.includes("Select a provider"), "the placeholder must not be sorted into the list");
  assert.ok(real.every((option) => option.disabled === false), "real providers are all selectable");
  assert.equal(provider.options.filter((option) => option.disabled).length, 1, "exactly one disabled option: the placeholder");
  assert.ok(labels.includes("Anthropic"), "models.dev providers stay merged in");
});

test("R3/R4.1 — empty state: endpoint value empty with https:// placeholder, model disabled, reasoning disabled at lowest", async () => {
  const pane = await bootPane();
  const endpoint = byId(pane, "aiEndpoint");
  assert.equal(endpoint.value, "", "endpoint value is empty on a fresh install");
  assert.equal(endpoint.placeholder, "https://", "endpoint placeholder is exactly https://");
  const model = byId(pane, "aiModelSelect");
  assert.equal(model.disabled, true, "model picker is disabled without a provider");
  assert.equal(model.options.length, 0, "model picker is empty without a provider");
  assert.equal(model.title, "Select a provider first", "model picker says what is missing");
  const reasoning = byId(pane, "aiReasoning");
  assert.equal(reasoning.disabled, true, "reasoning picker is disabled without a provider");
  assert.equal(reasoning.value, "lowest", "the stored sentinel stays lowest");
});

// ── 2. nothing rewrites the default or hits the network ─────────────────────
test("R2.1 — the empty state writes no provider and no OpenAI endpoint anywhere", async () => {
  const pane = await bootPane();
  const stored = pane.storage.has(SETTINGS_KEY) ? JSON.parse(pane.storage.get(SETTINGS_KEY)) : {};
  assert.ok(!("aiProvider" in stored) || stored.aiProvider === "", `stored aiProvider must stay empty, got ${JSON.stringify(stored.aiProvider)}`);
  assert.ok(!stored.aiEndpoint, `stored aiEndpoint must stay empty, got ${JSON.stringify(stored.aiEndpoint)}`);
  const rendered = JSON.stringify({
    providers: labelsOf(pane, "aiProvider"),
    endpoint: byId(pane, "aiEndpoint").value,
    endpointPlaceholder: byId(pane, "aiEndpoint").placeholder,
  });
  assert.ok(!rendered.includes("https://api.openai.com/v1"), "the empty state must not paint the OpenAI endpoint");
  assert.ok(!rendered.includes("gpt-"), "the empty state must not list OpenAI models");
});

test("R5.1 — a translation with an empty provider shows providerRequiredError and sends nothing", async () => {
  const pane = await bootPane();
  pane.resetCalls();
  const status = await pressTranslate(pane);
  assert.equal(status, REQUIRED, "translate reports the localized providerRequiredError");
  assert.equal(pane.calls.length, 0, `no request may leave the pane, saw ${pane.calls.map((call) => call.url).join(", ")}`);
  assert.ok(!byId(pane, "translate").disabled, "the button is re-enabled once the guard returns");
});

test("R5.1 — an AI entry with an empty provider sends no request even when the engine is AI", async () => {
  const pane = await bootPane({ storage: { [SETTINGS_KEY]: JSON.stringify({ engine: "ai", targetLanguage: "ja" }) } });
  assert.equal(byId(pane, "aiProvider").value, "", "a stored provider is not invented");
  pane.resetCalls();
  const status = await pressTranslate(pane);
  assert.equal(status, REQUIRED);
  assert.equal(pane.calls.length, 0);
});

test("R5.1 — the request path carries no openai fallback", () => {
  const script = fs.readFileSync(path.join(ROOT, "src", "taskpane.js"), "utf8");
  assert.doesNotMatch(script, /\|\|\s*["']openai["']/, 'no `|| "openai"` on a request path');
  assert.doesNotMatch(script, /\|\|\s*CUSTOM_PROVIDER/, "no `|| CUSTOM_PROVIDER` on a request path");
  assert.doesNotMatch(script, /aiProvider\s*\|\|\s*["']/, "the provider is never defaulted to a non-empty id");
});

// ── 3. leaving the empty state ──────────────────────────────────────────────
test("R3.1 — choosing a provider fills the endpoint value, loads models, enables reasoning", async () => {
  const pane = await bootPane();
  const provider = byId(pane, "aiProvider");
  assert.equal(provider.select("anthropic"), true, "a real provider can be picked");
  assert.equal(provider.value, "anthropic");
  assert.equal(byId(pane, "aiEndpoint").value, "https://api.anthropic.com/v1", "endpoint VALUE becomes the provider API base");
  assert.ok(optionOf(pane, "aiModelSelect", "") === undefined, "the model picker is no longer empty");
  const model = byId(pane, "aiModelSelect");
  assert.equal(model.disabled, false, "model picker enables");
  assert.ok(model.options.length > 1, `expected the anthropic models, got ${model.options.length}`);
  assert.equal(byId(pane, "aiReasoning").disabled, false, "reasoning enables");
  const placeholder = optionOf(pane, "aiProvider", "");
  assert.equal(placeholder.selected, false, "the placeholder is no longer the selection");
});

test("R2.1 — the empty state cannot be re-entered after a provider is chosen", async () => {
  const pane = await bootPane();
  // The placeholder stays in the list (still disabled, so the picker keeps
  // explaining itself) but is no longer the selection and cannot come back.
  assert.ok(optionOf(pane, "aiProvider", ""), "the placeholder is rendered on a fresh install");
  const provider = byId(pane, "aiProvider");
  provider.select("groq");
  const placeholder = optionOf(pane, "aiProvider", "");
  assert.ok(placeholder, "the disabled placeholder remains visible");
  assert.equal(placeholder.selected, false, "the placeholder is no longer the selection");
  assert.equal(placeholder.disabled, true, "the placeholder stays disabled");
  assert.equal(provider.select(""), false, "picking the placeholder is refused");
  assert.equal(provider.value, "groq", "the chosen provider stands");
  assert.ok(!labelsOf(pane, "aiProvider").some((label) => /^none$/i.test(label)), "there is no None item");
  assert.equal(byId(pane, "aiProvider").title, "", "the required-picker hint is gone once configured");
  // The rest of the pane agrees with the picker: nothing goes back to empty.
  assert.equal(byId(pane, "aiReasoning").disabled, false, "reasoning stays enabled");
  assert.equal(byId(pane, "aiModelSelect").disabled, false, "the model picker stays enabled");
});

test("R3.1 — a user-edited endpoint survives a later provider switch", async () => {
  const pane = await bootPane();
  const endpoint = byId(pane, "aiEndpoint");
  byId(pane, "aiProvider").select("openai");
  endpoint.fill("https://proxy.example/v1");
  byId(pane, "aiProvider").select("anthropic");
  assert.equal(endpoint.value, "https://proxy.example/v1", "the override is never clobbered");
  byId(pane, "aiProvider").select("deepseek");
  assert.equal(endpoint.value, "https://proxy.example/v1");
});

// ── 4. the conformance payload, read back out of the controls ───────────────
test("SPEC 2.1.0 — the payload carries both states from the real controls", async () => {
  const payload = await buildPayload();
  assert.equal(payload.spec, "2.1.0");
  assert.equal(payload.provider.current, "", "unconfigured payload: provider.current is empty");
  assert.deepEqual(payload.provider.placeholder, {
    value: "",
    label: "Select a provider",
    disabled: true,
    selected: true,
  });
  assert.equal(payload.provider.fallback, "", "no fallback provider");
  assert.equal(payload.model.disabled, true);
  assert.equal(payload.model.selectedProvider, "");
  assert.equal(payload.model.selectedModel, "");
  assert.deepEqual(payload.model.options, []);
  assert.equal(payload.endpoint.current, "");
  assert.equal(payload.endpoint.placeholder, "https://");
  assert.equal(payload.reasoning.disabled, true);
  assert.equal(payload.reasoning.default, "lowest");
  assert.ok(payload.configured, "the payload carries the configured (anthropic) state too");
  assert.equal(payload.configured.provider.current, "anthropic");
  assert.equal(payload.configured.provider.placeholder.selected, false);
  assert.equal(payload.configured.model.disabled, false);
  assert.equal(payload.configured.model.selectedProvider, "anthropic");
  assert.equal(payload.configured.endpoint.current, "https://api.anthropic.com/v1");
  assert.equal(payload.configured.reasoning.disabled, false);
});

test("SPEC 2.1.0 — the validator accepts the live payload under the 2.1 kit", async () => {
  const validateUrl = pathToFileURL(path.join(ROOT, "scripts", "conformance", "validate.mjs")).href;
  const { validatePayload, SPEC_VERSION } = await import(validateUrl);
  assert.equal(SPEC_VERSION, "2.1.0");
  const payload = await buildPayload();
  const summary = validatePayload(payload, {});
  assert.equal(summary.host, "dualtran-outlook");
});

test("SPEC 2.1.0 — the validator rejects the pre-2.1 shapes this change removed", async () => {
  const validateUrl = pathToFileURL(path.join(ROOT, "scripts", "conformance", "validate.mjs")).href;
  const { validatePayload, ConformanceError } = await import(validateUrl);
  const ruleOf = async (mutate) => {
    const payload = await buildPayload();
    validatePayload(payload, {});
    mutate(payload);
    try {
      validatePayload(payload, {});
    } catch (error) {
      assert.ok(error instanceof ConformanceError, "the validator must throw a ConformanceError");
      return error.rule;
    }
    return null;
  };
  // The default this task removes: a fresh install silently on openai.
  assert.equal(
    await ruleOf((payload) => {
      payload.provider.current = "openai";
    }),
    "R2",
    "an unconfigured placeholder with a provider selected is rejected"
  );
  assert.equal(
    await ruleOf((payload) => {
      payload.provider.fallback = "openai";
    }),
    "R2",
    "a silent openai fallback is rejected"
  );
  assert.equal(
    await ruleOf((payload) => {
      payload.endpoint.current = "https://api.openai.com/v1";
    }),
    "R3",
    "the OpenAI endpoint painted into the empty state is rejected"
  );
  assert.equal(
    await ruleOf((payload) => {
      payload.model.disabled = false;
    }),
    "R2",
    "an enabled model picker without a provider is rejected"
  );
});

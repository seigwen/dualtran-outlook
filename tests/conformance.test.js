"use strict";
/**
 * Host conformance for the Outlook task pane (SPEC 2.0.0, node:test).
 *
 * Three layers, no jsdom and no network:
 *   1. markup      — src/taskpane.html really renders <select id="target">;
 *   2. behaviour   — the real src/catalog.js + src/languages.js answer the
 *                    questions the payload asks (endpoint rule, one merged
 *                    provider list, ≥40 human-readable languages);
 *   3. spec        — the kit's own selfTest() and validatePayload() over the
 *                    payload built from those two modules.
 */
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const ROOT = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(ROOT, "src", "taskpane.html"), "utf8");
const validatorUrl = pathToFileURL(path.join(ROOT, "scripts", "conformance", "validate.mjs")).href;
const payloadUrl = pathToFileURL(path.join(ROOT, "scripts", "conformance", "payload.mjs")).href;

const readTag = (markup, id) => {
  const pattern = new RegExp(`<([a-z0-9-]+)\\b[^>]*\\bid="${id}"[^>]*>`, "i");
  const match = markup.match(pattern);
  return match ? { tag: match[1].toLowerCase(), text: match[0] } : null;
};

test("R1 — the target control is a <select>, not a text input", () => {
  const control = readTag(html, "target");
  assert.ok(control, "taskpane.html has no element with id=\"target\"");
  assert.equal(control.tag, "select", `#target must be a <select>, got <${control.tag}>`);
  assert.doesNotMatch(control.text, /\btype="text"/i, "#target must not carry a text-input attribute");
  assert.doesNotMatch(control.text, /\bsize=/, "a <select> has no size attribute for a one-line picker");
  const source = readTag(html, "source");
  assert.ok(source && source.tag === "input", "the source language stays a free-text field (it accepts \"auto\")");
});

test("R1 — languages.js offers ≥40 human-readable targets and keeps the current one", () => {
  const languages = require(path.join(ROOT, "src", "languages.js"));
  const list = languages.TARGET_LANGUAGES;
  assert.ok(Array.isArray(list), "TARGET_LANGUAGES must be an array");
  assert.ok(list.length >= 40, `expected ≥40 target languages, got ${list.length}`);
  const names = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "conformance", "language-names.json"), "utf8"));
  for (const entry of list) {
    assert.equal(typeof entry.code, "string", "every entry needs a code");
    assert.equal(entry.code, entry.code.trim(), `code "${entry.code}" carries whitespace`);
    assert.notEqual(entry.code.toLowerCase(), "auto", "a target list cannot offer auto");
    assert.ok(names[entry.code], `"${entry.code}" is not in tests/fixtures/conformance/language-names.json`);
    assert.equal(typeof entry.name, "string", `"${entry.code}" has no name`);
    assert.ok(entry.name.trim(), `"${entry.code}" has an empty name`);
    assert.equal(entry.label, `${entry.name} (${entry.code})`, `"${entry.code}" label must read "Name (code)"`);
  }
  for (const [code, name] of Object.entries(names)) {
    const entry = list.find((candidate) => candidate.code === code);
    assert.ok(entry, `tests/fixtures/conformance/language-names.json lists "${code}" but languages.js does not`);
    assert.equal(entry.name, name, `"${code}" must use the fact table's name, got "${entry.name}"`);
  }
  const codes = list.map((entry) => entry.code);
  assert.equal(new Set(codes).size, codes.length, "duplicate language codes");
  assert.ok(codes.includes("zh-CN"), "zh-CN is the default target and must be selectable");
  assert.equal(languages.DEFAULT_TARGET, "zh-CN");
  assert.ok(codes.includes(languages.DEFAULT_TARGET), "DEFAULT_TARGET must be in the list");
  assert.equal(languages.labelFor("ja"), "Japanese (ja)");
});

test("R2 — the provider list drops the divider and merges sources by visible label", () => {
  const catalog = require(path.join(ROOT, "src", "catalog.js"));
  const sample = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "conformance", "models-dev-sample.json"), "utf8"));
  const { providers, modelsById } = catalog.projectCatalog(sample);
  const list = catalog.buildProviderList(providers, modelsById);

  const compare = new Intl.Collator("en", { numeric: true, sensitivity: "base" });
  for (let index = 1; index < list.length; index += 1) {
    assert.ok(
      compare.compare(list[index - 1].name, list[index].name) <= 0,
      `"${list[index - 1].name}" must not be listed before "${list[index].name}"`
    );
  }
  assert.ok(list.some((entry) => entry.builtIn), "built-ins must stay in the list");
  assert.ok(list.some((entry) => !entry.builtIn), "models.dev entries must be merged in");
  const ids = list.map((entry) => entry.id);
  assert.equal(new Set(ids).size, ids.length, "duplicate provider ids");
  // The sample aliases moonshotai for the Moonshot built-in: exactly one survives.
  assert.equal(ids.filter((id) => id === "moonshot" || id === "moonshotai").length, 1);
  assert.ok(ids.includes("lmstudio"), "a models.dev-only provider must be reachable");
  // Every entry still resolves to a usable endpoint and a model.
  for (const entry of list) {
    assert.match(entry.api, /^https?:\/\//, `${entry.id} has no endpoint`);
    assert.ok(entry.models.length > 0, `${entry.id} has no models`);
  }
  // Built-ins keep their curated endpoint (models.dev carries none for them).
  assert.equal(list.find((entry) => entry.id === "openai").api, "https://api.openai.com/v1");
});

test("R2 — taskpane.js renders no divider and still creates the custom model entry", () => {
  const script = fs.readFileSync(path.join(ROOT, "src", "taskpane.js"), "utf8");
  assert.doesNotMatch(script, /__more__/, "the __more__ divider sentinel must be gone from the renderer");
  assert.doesNotMatch(script, /providerDivider/, "the divider label must not be painted");
  // A disabled <option> is what the old divider was; buttons may still disable.
  assert.doesNotMatch(script, /option\.disabled\s*=\s*true|divider\s*\.\s*disabled\s*=\s*true|\.disabled\s*=\s*true;\s*el\.aiProvider/, "no picker option may be disabled");
  const customAt = script.indexOf("custom.value = CUSTOM_VALUE");
  const appendAt = script.indexOf("el.aiModelSelect.appendChild(custom)");
  assert.ok(customAt >= 0 && appendAt > customAt, "the custom model option must still be appended to the model select");
});

test("R3 — resolveEndpointOnProviderChange fills an empty endpoint and keeps a user override", () => {
  const catalog = require(path.join(ROOT, "src", "catalog.js"));
  assert.equal(typeof catalog.resolveEndpointOnProviderChange, "function", "src/catalog.js must export the pure helper");
  const resolve = catalog.resolveEndpointOnProviderChange;
  assert.equal(resolve("", "https://api.openai.com/v1"), "https://api.openai.com/v1");
  assert.equal(resolve("   ", "https://api.deepseek.com/v1"), "https://api.deepseek.com/v1", "blank counts as empty");
  assert.equal(resolve(undefined, "http://localhost:11434/v1"), "http://localhost:11434/v1");
  assert.equal(resolve(null, "https://api.x.ai/v1"), "https://api.x.ai/v1");
  const user = "https://proxy.example/v1";
  assert.equal(resolve(user, "https://api.openai.com/v1"), user, "a user-edited endpoint must survive");
  assert.equal(resolve(`  ${user}  `, "https://api.anthropic.com/v1"), user, "the override is trimmed, not replaced");
  assert.equal(resolve("https://proxy.example/v1", ""), user, "a provider without an api changes nothing");
  assert.equal(resolve("", ""), "", "nothing to fill with");
});

test("R5 — the pane never hard-codes the divider copy", () => {
  const script = fs.readFileSync(path.join(ROOT, "src", "taskpane.js"), "utf8");
  for (const match of script.matchAll(/\bt\(\s*"([A-Za-z0-9_]+)"/g)) {
    assert.notEqual(match[1], "providerDivider", "providerDivider is no longer rendered");
  }
});

test("R2 — the model sort the pane renders matches the validator's collation", () => {
  // The pane sorts option text with localeCompare; the kit compares labels with
  // Intl.Collator({numeric:true}). If the flags drift apart ("GPT OSS 120B" vs
  // "GPT OSS 20B" then order differently) the picker shows an order conformance
  // rejects. Read the flags out of taskpane.js so this test tracks the real code.
  const script = fs.readFileSync(path.join(ROOT, "src", "taskpane.js"), "utf8");
  const flags = script.match(/\.localeCompare\([^,]+,\s*"en",\s*(\{[^}]*\})\)/);
  assert.ok(flags, "taskpane.js no longer sorts option text with a localeCompare flags argument");
  const paneOpts = Function(`return (${flags[1]});`)();

  const catalog = require(path.join(ROOT, "src", "catalog.js"));
  const sample = JSON.parse(fs.readFileSync(path.join(ROOT, "tests", "fixtures", "conformance", "models-dev-sample.json"), "utf8"));
  const { providers, modelsById } = catalog.projectCatalog(sample);
  const list = catalog.buildProviderList(providers, modelsById);
  const validatorCompare = new Intl.Collator("en", { numeric: true, sensitivity: "base" }).compare;
  const paneSort = (a, b) => a.localeCompare(b, "en", paneOpts);
  for (const entry of list) {
    const catalogue = entry.models.map((model) => model.label).sort(paneSort);
    for (let index = 1; index < catalogue.length; index += 1) {
      assert.ok(
        validatorCompare(catalogue[index - 1], catalogue[index]) <= 0,
        `${entry.id}: pane would render "${catalogue[index - 1]}" before "${catalogue[index]}"`
      );
    }
  }
  // The free-text action is appended after the sorted catalogue, never sorted
  // into it (SPEC R2: custom model entries are last).
  const paneScript = fs.readFileSync(path.join(ROOT, "src", "taskpane.js"), "utf8");
  assert.ok(
    paneScript.indexOf("el.aiModelSelect.appendChild(custom)") > paneScript.indexOf("for (const model of models)"),
    "the custom model option must be appended after the catalogue loop"
  );
  assert.doesNotMatch(paneScript, /sortOptionsByLabel/, "the custom entry must not be sorted into the catalogue");
});

test("kit self-test catches its own negative controls", async () => {
  const { selfTest } = await import(validatorUrl);
  assert.equal(selfTest(), 12);
});

test("validatePayload accepts the payload built from the real modules", async () => {
  const { validatePayload } = await import(validatorUrl);
  const { buildPayload } = await import(payloadUrl);
  const payload = buildPayload();
  assert.equal(payload.host, "dualtran-outlook");
  const summary = validatePayload(payload, {});
  assert.equal(summary.host, "dualtran-outlook");
  assert.ok(summary.languages >= 40, `payload carries ${summary.languages} languages`);
  assert.ok(summary.providers > 11, `payload carries ${summary.providers} providers`);
  assert.ok(summary.localeCount >= 38, `payload carries ${summary.localeCount} locales`);
});

test("validatePayload rejects the pre-fix shape (text target, divider)", async () => {
  const { validatePayload, ConformanceError } = await import(validatorUrl);
  const { buildPayload } = await import(payloadUrl);
  // Each mutation starts from a passing payload: one broken rule at a time.
  const ruleOf = (mutate) => {
    const payload = buildPayload();
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
  assert.equal(ruleOf((payload) => { payload.language.control = "text"; }), "R1");
  assert.equal(
    ruleOf((payload) => {
      payload.provider.options.splice(3, 0, { value: "__more__", label: "──── more ────", disabled: true });
    }),
    "R2"
  );
  assert.equal(ruleOf((payload) => { payload.endpoint.autofill.afterSwitch = "https://api.openai.com/v1"; }), "R3");
});

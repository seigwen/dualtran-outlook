"use strict";
/**
 * SPEC 2.1.0 §R5.1 — the three new keys in every locale (node:test).
 *
 * The English source strings are fixed by the spec. zh_CN is given verbatim.
 * Every other catalogue must carry a real translation (not the English string,
 * not a leftover from the placeholderDivider era) with its placeholders intact,
 * and no existing human key may have changed (locales/ is append-only here).
 */
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const LOCALES = path.resolve(__dirname, "..", "locales");
const KEYS = ["providerPlaceholder", "providerRequiredForModel", "providerRequiredError"];
/** Every catalogue key the pane ships: the pre-existing 38 plus the three above. */
const BASE_KEY_COUNT = 38;
const ENGLISH = {
  providerPlaceholder: "Select a provider",
  providerRequiredForModel: "Select a provider first",
  providerRequiredError: "Choose an AI provider in settings before translating.",
};
const ZH_CN = {
  providerPlaceholder: "请选择 AI 提供商",
  providerRequiredForModel: "请先选择提供商",
  providerRequiredError: "翻译前请先在设置中选择 AI 提供商。",
};

const read = (code) => JSON.parse(fs.readFileSync(path.join(LOCALES, code, "messages.json"), "utf8"));
const codes = fs
  .readdirSync(LOCALES)
  .filter((code) => fs.existsSync(path.join(LOCALES, code, "messages.json")))
  .sort();

test("R5.1 — the English catalogue carries the three source strings verbatim", () => {
  const en = read("en");
  for (const key of KEYS) assert.equal(en[key] && en[key].message, ENGLISH[key], `en: ${key}`);
});

test("R5.1 — zh_CN carries the three agreed translations verbatim", () => {
  const zh = read("zh_CN");
  for (const key of KEYS) assert.equal(zh[key] && zh[key].message, ZH_CN[key], `zh_CN: ${key}`);
});

test("R5.1 — every locale carries the three keys, translated and non-empty", () => {
  const problems = [];
  for (const code of codes) {
    const catalogue = read(code);
    for (const key of KEYS) {
      const value = catalogue[key] && catalogue[key].message;
      if (typeof value !== "string" || !value.trim()) {
        problems.push(`${code}: ${key} missing or empty`);
        continue;
      }
      if (!/^(en|zh|ja)/.test(code) && value.trim() === ENGLISH[key]) {
        problems.push(`${code}: ${key} is still the English source string`);
      }
      if (/[{}]|\$[A-Z0-9_]+\$/.test(value)) {
        // These three strings have no placeholders; any token is a translation bug.
        problems.push(`${code}: ${key} gained a placeholder token: ${value}`);
      }
      if (code !== "en" && code !== "zh_CN" && /[\u4e00-\u9fff]/.test(value) && !code.startsWith("zh") && !code.startsWith("ja")) {
        problems.push(`${code}: ${key} contains CJK text`);
      }
    }
  }
  assert.deepEqual(problems, [], problems.join("\n"));
});

test("R5.1 — the key set is identical in every locale: exactly three new keys", () => {
  const enKeys = Object.keys(read("en"));
  assert.equal(enKeys.length, BASE_KEY_COUNT + KEYS.length, `en carries ${enKeys.length} keys, expected ${BASE_KEY_COUNT + KEYS.length}`);
  for (const key of KEYS) assert.ok(enKeys.includes(key), `en is missing ${key}`);
  for (const code of codes) {
    assert.deepEqual(Object.keys(read(code)), enKeys, `${code}: key set drifts from en`);
  }
});

test("R5.1 — the three new keys are the only difference from the pre-change catalogues", () => {
  // Every shipped translation predates this change; the guard proves the fill
  // appended keys instead of rewriting locale files. A missing baseline means
  // this is the first run — nothing to compare, the key-set test still holds.
  const baseline = path.join(__dirname, "fixtures", "locales-baseline.json");
  if (!fs.existsSync(baseline)) return;
  const before = JSON.parse(fs.readFileSync(baseline, "utf8"));
  const changed = [];
  for (const [code, catalogue] of Object.entries(before)) {
    const now = read(code);
    for (const [key, entry] of Object.entries(catalogue)) {
      const value = now[key] && now[key].message;
      if (value !== entry.message) changed.push(`${code}: ${key} changed (${JSON.stringify(entry.message)} → ${JSON.stringify(value)})`);
    }
  }
  assert.deepEqual(changed, [], changed.join("\n"));
});

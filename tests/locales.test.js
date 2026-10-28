"use strict";
/**
 * Locale guards for the Outlook add-in (node:test).
 *
 * The task pane ships the same 38 languages as the browser extension; the
 * manifest's add-in description is localised through <Override Locale>, which
 * scripts/build-manifest-i18n.mjs regenerates from these catalogues.
 */
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const LOCALES = path.join(ROOT, "locales");
const EXPECTED = [
  "af", "ar", "ca", "cs", "da", "de", "el", "en", "es", "fa", "fi", "fr", "he", "hi", "hr",
  "hu", "is", "it", "ja", "ko", "lv", "nl", "no", "pl", "pt_BR", "pt_PT", "ro", "ru", "sl",
  "sr", "sv", "th", "tr", "ug", "uk", "vi", "zh_CN", "zh_TW",
];
const CULTURE = {
  af: "af-ZA", ar: "ar-SA", ca: "ca-ES", cs: "cs-CZ", da: "da-DK", de: "de-DE", el: "el-GR",
  en: "en-US", es: "es-ES", fa: "fa-IR", fi: "fi-FI", fr: "fr-FR", he: "he-IL", hi: "hi-IN",
  hr: "hr-HR", hu: "hu-HU", is: "is-IS", it: "it-IT", ja: "ja-JP", ko: "ko-KR", lv: "lv-LV",
  nl: "nl-NL", no: "nb-NO", pl: "pl-PL", pt_BR: "pt-BR", pt_PT: "pt-PT", ro: "ro-RO",
  ru: "ru-RU", sl: "sl-SI", sr: "sr-Latn-RS", sv: "sv-SE", th: "th-TH", tr: "tr-TR",
  ug: "ug-CN", uk: "uk-UA", vi: "vi-VN", zh_CN: "zh-CN", zh_TW: "zh-TW",
};

const read = (code) => JSON.parse(fs.readFileSync(path.join(LOCALES, code, "messages.json"), "utf8"));
const en = read("en");
const enKeys = Object.keys(en);

test("ships the agreed 38 locales", () => {
  const codes = fs.readdirSync(LOCALES).filter((code) => fs.existsSync(path.join(LOCALES, code, "messages.json")));
  assert.deepEqual(codes.sort(), [...EXPECTED].sort());
});

test("every locale carries every key, non-empty, without stray Chinese", () => {
  for (const code of EXPECTED) {
    const catalogue = read(code);
    assert.deepEqual(Object.keys(catalogue), enKeys, `${code}: key set`);
    for (const key of enKeys) {
      const value = catalogue[key] && catalogue[key].message;
      assert.equal(typeof value === "string" && value.trim().length > 0, true, `${code}: ${key}`);
    }
    if (!code.startsWith("zh") && !code.startsWith("ja")) {
      assert.equal(/[\u4e00-\u9fff]/.test(JSON.stringify(catalogue)), false, `${code}: Chinese left over`);
    }
  }
});

test("no user-visible literal bypasses the catalogue", () => {
  const js = fs.readFileSync(path.join(ROOT, "src", "taskpane.js"), "utf8");
  const html = fs.readFileSync(path.join(ROOT, "src", "taskpane.html"), "utf8");
  const offenders = [];
  for (const pattern of [/status\(\s*"[^"]/g, /\.textContent\s*=\s*"[^"]/g, /setAttribute\(\s*"placeholder",\s*"[^"]/g]) {
    for (const match of js.matchAll(pattern)) offenders.push(`taskpane.js: ${match[0]}`);
  }
  for (const tag of html.match(/<[^>]+>/g) || []) {
    if (/placeholder="/.test(tag) && !/data-i18n-placeholder=/.test(tag) && !/placeholder="(?:sk-…|https:\/\/)/.test(tag)) {
      offenders.push(`taskpane.html: ${tag.trim().slice(0, 70)}`);
    }
  }
  assert.deepEqual(offenders, [], offenders.join("\n"));
});

test("the bundled data file matches the catalogues", () => {
  const data = fs.readFileSync(path.join(ROOT, "src", "i18n.data.js"), "utf8");
  for (const code of EXPECTED) assert.ok(data.includes(`"${code}":`), `i18n.data.js is missing ${code}`);
});

test("the manifest carries a description override per culture", () => {
  const manifest = fs.readFileSync(path.join(ROOT, "manifest.xml"), "utf8");
  for (const code of EXPECTED) {
    const culture = CULTURE[code];
    if (code === "en") {
      // English is the DefaultValue, not an Override.
      assert.ok(manifest.includes('<Description DefaultValue="'), "manifest.xml has no Description DefaultValue");
      continue;
    }
    assert.ok(
      manifest.includes(`<Override Locale="${culture}"`),
      `manifest.xml has no Description override for ${culture}`
    );
  }
});

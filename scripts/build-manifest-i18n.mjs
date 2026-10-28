#!/usr/bin/env node
/**
 * build-manifest-i18n.mjs — rewrite the add-in <Description> with one
 * <Override Locale="…"> per shipped language (Office picks the user's culture;
 * en-US stays the DefaultValue).
 *
 * Idempotent: the whole element is regenerated from the locale catalogues.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const LOCALES = path.join(ROOT, "locales");
const MANIFEST = path.join(ROOT, "manifest.xml");

/** Our catalogue code → Office culture name. */
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
const defaultText = en.extDescription.message;

const escape = (value) => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const overrides = Object.keys(CULTURE)
  .filter((code) => code !== "en")
  .map((code) => `    <Override Locale="${CULTURE[code]}" Value="${escape(read(code).extDescription.message)}"/>`)
  .join("\n");

const element = `<Description DefaultValue="${escape(defaultText)}">\n${overrides}\n  </Description>`;

let manifest = fs.readFileSync(MANIFEST, "utf8");
if (!/<Description[\s\S]*?<\/Description>/.test(manifest)) {
  manifest = manifest.replace(/<Description DefaultValue="[^"]*"\s*\/>/, element);
} else {
  manifest = manifest.replace(/<Description[\s\S]*?<\/Description>/, element);
}
fs.writeFileSync(MANIFEST, manifest);
console.log(`[build-manifest-i18n] Description + ${Object.keys(CULTURE).length - 1} Override Locale entries`);

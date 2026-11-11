/**
 * languages.js — target-language table for the Outlook task pane.
 *
 * The pane's header used to be a free-text <input size="4">, which let a typo
 * reach Google/AI as a bogus code. The header is now a <select> fed from here.
 *
 * Names come from `facts/language-names.json` (the shared table every DualTran
 * host uses); the list is the same 38 cultures the locales/ directory ships,
 * plus the regional variants the translation engines actually distinguish.
 * Labels read "Name (code)" so the code stays visible while the name stays
 * readable. `auto` is a *source*-language concept and is deliberately absent.
 *
 * Exposed as `window.DualTranLanguages` and as `module.exports`, same UMD shape
 * as catalog.js, so the Node tests can require it without a DOM.
 */
(function (root, factory) {
  "use strict";
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.DualTranLanguages = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  /** The pane's default target; also the fallback when a stored code is unknown. */
  const DEFAULT_TARGET = "zh-CN";

  /** [code, name] — name is the English exonym, codes are BCP 47 / engine codes. */
  const ENTRIES = [
    ["af", "Afrikaans"],
    ["ar", "Arabic"],
    ["bg", "Bulgarian"],
    ["bn", "Bengali"],
    ["ca", "Catalan"],
    ["cs", "Czech"],
    ["da", "Danish"],
    ["de", "German"],
    ["el", "Greek"],
    ["en", "English"],
    ["es", "Spanish"],
    ["et", "Estonian"],
    ["fa", "Persian"],
    ["fi", "Finnish"],
    ["fr", "French"],
    ["he", "Hebrew"],
    ["hi", "Hindi"],
    ["hr", "Croatian"],
    ["hu", "Hungarian"],
    ["id", "Indonesian"],
    ["it", "Italian"],
    ["ja", "Japanese"],
    ["ko", "Korean"],
    ["lt", "Lithuanian"],
    ["lv", "Latvian"],
    ["ms", "Malay"],
    ["nl", "Dutch"],
    ["no", "Norwegian"],
    ["pl", "Polish"],
    ["pt", "Portuguese"],
    ["pt-BR", "Portuguese (Brazil)"],
    ["ro", "Romanian"],
    ["ru", "Russian"],
    ["sk", "Slovak"],
    ["sl", "Slovenian"],
    ["sr", "Serbian"],
    ["sv", "Swedish"],
    ["th", "Thai"],
    ["tr", "Turkish"],
    ["uk", "Ukrainian"],
    ["ur", "Urdu"],
    ["vi", "Vietnamese"],
    ["zh-CN", "Chinese (Simplified)"],
    ["zh-TW", "Chinese (Traditional)"],
  ];

  const collate = new Intl.Collator("en", { sensitivity: "base" });

  /**
   * Targets sorted by their visible name — the <select> needs no second sort,
   * and "Chinese (Simplified)" sits next to "Chinese (Traditional)".
   */
  const TARGET_LANGUAGES = ENTRIES.map(([code, name]) => ({ code, name, label: `${name} (${code})` })).sort(
    (a, b) => collate.compare(a.name, b.name) || collate.compare(a.code, b.code)
  );

  const BY_CODE = new Map(TARGET_LANGUAGES.map((entry) => [entry.code.toLowerCase(), entry]));

  /** The code to keep in the <select>: the stored one when known, else the default. */
  function normalizeTarget(code) {
    const wanted = String(code || "").trim().toLowerCase();
    const found = BY_CODE.get(wanted);
    return found ? found.code : DEFAULT_TARGET;
  }

  /** Display label for a code, for status text or titles. */
  function labelFor(code) {
    const found = BY_CODE.get(String(code || "").trim().toLowerCase());
    return found ? found.label : "";
  }

  return { DEFAULT_TARGET, TARGET_LANGUAGES, normalizeTarget, labelFor };
});

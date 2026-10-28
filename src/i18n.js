/**
 * Tiny i18n runtime for the Outlook task pane.
 *
 * Catalogues live in `locales/<code>/messages.json` (same set as the browser
 * extension) and are bundled into `src/i18n.data.js` by scripts/build-i18n.mjs —
 * the pane loads plain <script> files, so no fetch at runtime.
 */
(function (root) {
  "use strict";

  function normalizeTag(tag) {
    if (!tag || typeof tag !== "string") return "";
    var raw = tag.trim().replace(/_/g, "-").toLowerCase();
    if (!raw) return "";
    if (raw.indexOf("zh-hant") === 0 || raw === "zh-tw" || raw === "zh-hk" || raw === "zh-mo") return "zh_TW";
    if (raw.indexOf("zh-hans") === 0 || raw.indexOf("zh-cn") === 0 || raw === "zh" || raw === "zh-sg") return "zh_CN";
    if (raw.indexOf("zh-") === 0) return "zh_CN";
    if (raw === "pt-br") return "pt_BR";
    if (raw === "pt-pt" || raw === "pt") return "pt_PT";
    if (raw === "nb" || raw === "nb-no" || raw === "nn" || raw === "no-no") return "no";
    return raw.split("-")[0];
  }

  function detectLocale(preferred, available) {
    var codes = available && available.length ? available : ["en"];
    var exact = normalizeTag(preferred);
    if (!exact) return codes.indexOf("en") >= 0 ? "en" : codes[0];
    var direct = codes.filter(function (code) {
      return String(code).toLowerCase() === exact.toLowerCase();
    })[0];
    if (direct) return direct;
    var baseTag = String(preferred).trim().replace(/_/g, "-").toLowerCase().split("-")[0];
    var base = codes.filter(function (code) {
      return String(code).toLowerCase().split("_")[0] === baseTag;
    })[0];
    if (base) return base;
    return codes.indexOf("en") >= 0 ? "en" : codes[0];
  }

  function interpolate(text, subs) {
    if (!subs) return text;
    return String(text).replace(/\$([A-Z0-9_]+)\$/g, function (whole, name) {
      return Object.prototype.hasOwnProperty.call(subs, name) ? String(subs[name]) : whole;
    });
  }

  function create(locales, preferred) {
    var codes = Object.keys(locales || {});
    var code = detectLocale(preferred, codes);
    var table = (locales && locales[code]) || {};
    var fallback = (locales && locales.en) || {};
    function t(key, subs) {
      var entry = table[key] || fallback[key];
      var text = entry && typeof entry === "object" ? entry.message : entry;
      if (typeof text !== "string") return key;
      return interpolate(text, subs);
    }
    return { code: code, t: t };
  }

  function apply(rootNode, t) {
    if (!rootNode || !rootNode.querySelectorAll) return;
    var texts = rootNode.querySelectorAll("[data-i18n]");
    for (var i = 0; i < texts.length; i++) {
      var key = texts[i].getAttribute("data-i18n");
      if (key) texts[i].textContent = t(key);
    }
    var placeholders = rootNode.querySelectorAll("[data-i18n-placeholder]");
    for (var j = 0; j < placeholders.length; j++) {
      var pkey = placeholders[j].getAttribute("data-i18n-placeholder");
      if (pkey) placeholders[j].setAttribute("placeholder", t(pkey));
    }
    var titles = rootNode.querySelectorAll("[data-i18n-title]");
    for (var k = 0; k < titles.length; k++) {
      var tkey = titles[k].getAttribute("data-i18n-title");
      if (tkey) titles[k].setAttribute("title", t(tkey));
    }
  }

  /** Office tells us the user's display language; fall back to the browser. */
  function displayLocale() {
    try {
      if (root.Office && root.Office.context && root.Office.context.displayLanguage) {
        return root.Office.context.displayLanguage;
      }
    } catch (error) {
      /* Office.js not ready */
    }
    return root.navigator && root.navigator.language ? root.navigator.language : "en";
  }

  var RTL = { ar: 1, fa: 1, he: 1, ug: 1 };

  /** Mirror the pane for right-to-left scripts and keep <html lang> honest. */
  function applyDirection(i18n) {
    var doc = root.document;
    if (!doc || !doc.documentElement) return;
    if (RTL[i18n.code]) doc.documentElement.setAttribute("dir", "rtl");
    else doc.documentElement.removeAttribute("dir");
    doc.documentElement.setAttribute("lang", String(i18n.code).replace("_", "-"));
  }

  function init(preferred) {
    var i18n = create(root.DT_LOCALES || {}, preferred || displayLocale());
    i18n.apply = function (node) {
      apply(node || root.document, i18n.t);
      applyDirection(i18n);
    };
    root.DT_I18N = i18n;
    return i18n;
  }

  root.DtI18n = { normalizeTag: normalizeTag, detectLocale: detectLocale, create: create, apply: apply, init: init };
})(typeof window !== "undefined" ? window : this);

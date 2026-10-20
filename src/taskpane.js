/**
 * taskpane.js — Office.js glue: read the selected text or the whole message,
 * translate it, and optionally write the result back into the draft.
 */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const el = {
    engine: $("engine"), source: $("source"), target: $("target"),
    input: $("input"), translate: $("translate"), load: $("load"),
    copy: $("copy"), status: $("status"), result: $("result"),
    replace: $("replace"), insert: $("insert"),
    settings: $("settings"), panel: $("settingsPanel"),
    aiEndpoint: $("aiEndpoint"), aiModel: $("aiModel"), aiKey: $("aiKey"), save: $("saveSettings"),
  };

  const SETTINGS_KEY = "dualtran.outlook.settings";

  function loadSettings() {
    try { return JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}"); } catch { return {}; }
  }
  function saveSettings(patch) {
    const next = Object.assign(loadSettings(), patch);
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(next));
    return next;
  }
  function paintSettings() {
    const s = loadSettings();
    el.engine.value = s.engine || "google";
    el.source.value = s.sourceLanguage || "auto";
    el.target.value = s.targetLanguage || "zh-CN";
    el.aiEndpoint.value = s.aiEndpoint || "";
    el.aiModel.value = s.aiModel || "";
    el.aiKey.value = s.aiKey || "";
  }
  function current() {
    return {
      engine: el.engine.value,
      sourceLanguage: el.source.value.trim() || "auto",
      targetLanguage: el.target.value.trim() || "en",
      aiEndpoint: el.aiEndpoint.value.trim(),
      aiModel: el.aiModel.value.trim(),
      aiKey: el.aiKey.value.trim(),
    };
  }
  function status(msg, isError) {
    el.status.textContent = msg || "";
    el.status.classList.toggle("error", !!isError);
  }

  function withItem(fn) {
    const item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
    if (!item) { status("Open a message or a draft first.", true); return; }
    fn(item);
  }

  function getSelected() {
    return new Promise((resolve) => {
      withItem((item) => {
        if (item.getSelectedDataAsync) {
          item.getSelectedDataAsync(Office.CoercionType.Text, (r) => {
            resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value && r.value.data) || "" : "");
          });
        } else resolve("");
      });
    });
  }

  function getBody() {
    return new Promise((resolve, reject) => {
      withItem((item) => {
        item.body.getAsync(Office.CoercionType.Text, (r) => {
          if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value || "");
          else reject(new Error(r.error && r.error.message || "body read failed"));
        });
      });
    });
  }

  async function translateNow() {
    const text = el.input.value.trim();
    if (!text) { status("Nothing to translate.", true); return; }
    const cfg = current();
    saveSettings(cfg);
    el.translate.disabled = true;
    status("Translating…");
    el.result.textContent = "";
    el.copy.disabled = true;
    el.replace.disabled = true;
    el.insert.disabled = true;
    try {
      const out = cfg.engine === "ai"
        ? await DualTranTranslate.ai(text, cfg.sourceLanguage, cfg.targetLanguage, cfg)
        : await DualTranTranslate.google(text, cfg.sourceLanguage, cfg.targetLanguage);
      el.result.textContent = out;
      el.copy.disabled = !out;
      el.replace.disabled = !out;
      el.insert.disabled = !out;
      status(out ? "Done." : "Empty translation.");
    } catch (e) {
      status(String(e && e.message || e), true);
    } finally {
      el.translate.disabled = false;
    }
  }

  function writeBody(mode) {
    const out = el.result.textContent;
    if (!out) return;
    withItem((item) => {
      const cb = (r) => status(r.status === Office.AsyncResultStatus.Succeeded ? "Written to the message." : "Write failed.", r.status !== Office.AsyncResultStatus.Succeeded);
      if (mode === "replace") item.body.setAsync(out, { coercionType: Office.CoercionType.Text }, cb);
      else item.body.setSelectedDataAsync(out, { coercionType: Office.CoercionType.Text }, cb);
    });
  }

  function wire() {
    el.translate.addEventListener("click", translateNow);
    el.copy.addEventListener("click", () => navigator.clipboard.writeText(el.result.textContent).then(() => status("Copied.")));
    el.load.addEventListener("click", async () => {
      const selected = await getSelected();
      if (selected) { el.input.value = selected; status("Loaded the selection."); return; }
      try { el.input.value = await getBody(); status("Loaded the whole message."); }
      catch (e) { status(String(e.message || e), true); }
    });
    el.replace.addEventListener("click", () => writeBody("replace"));
    el.insert.addEventListener("click", () => writeBody("insert"));
    el.settings.addEventListener("click", () => { el.panel.hidden = !el.panel.hidden; });
    el.save.addEventListener("click", () => { saveSettings(current()); status("Saved."); });
    el.input.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) translateNow();
    });
  }

  if (typeof Office !== "undefined") {
    Office.onReady(() => { paintSettings(); wire(); });
  } else {
    // plain-browser preview (no Office host): still usable for a quick check
    document.addEventListener("DOMContentLoaded", () => { paintSettings(); wire(); status("Preview mode — Office.js not available."); });
  }
})();

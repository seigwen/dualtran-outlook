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

  el.aiProvider = $("aiProvider");
  el.aiModelSelect = $("aiModelSelect");
  el.aiReasoning = $("aiReasoning");
  el.customModelRow = $("customModelRow");

  const SETTINGS_KEY = "dualtran.outlook.settings";
  const CATALOG_KEY = "dualtran.outlook.catalog";
  const CATALOG_TTL = 24 * 60 * 60 * 1000;
  const MORE_VALUE = "__more__";
  const CUSTOM_VALUE = "__custom__";

  /** Projected models.dev data: { providers, modelsById, reasoning }. */
  let catalog = { providers: [], modelsById: {}, reasoning: {} };
  /** Built-ins first, then the catalogue — what the dropdown shows. */
  let providerList = [];

  function readCatalogCache() {
    try { return JSON.parse(localStorage.getItem(CATALOG_KEY) || "null"); } catch { return null; }
  }

  /**
   * models.dev, cached for a day. A failed refresh keeps the stale copy so a
   * flaky network never empties the provider list.
   */
  async function loadCatalog() {
    const cached = readCatalogCache();
    if (cached && Date.now() - cached.ts < CATALOG_TTL) { catalog = cached.data; return; }
    try {
      const res = await fetch(DualTranCatalog.MODELS_DEV_URL);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = DualTranCatalog.projectCatalog(await res.json());
      localStorage.setItem(CATALOG_KEY, JSON.stringify({ ts: Date.now(), data }));
      catalog = data;
    } catch (e) {
      catalog = cached ? cached.data : { providers: [], modelsById: {}, reasoning: {} };
    }
  }

  function providerById(id) {
    return providerList.find((p) => p.id === id);
  }

  function devIdFor(providerId) {
    const builtIn = DualTranCatalog.BUILTINS.find((entry) => entry.id === providerId);
    return builtIn ? builtIn.dev : providerId;
  }

  function selectedModel() {
    if (el.aiModelSelect.value === CUSTOM_VALUE) return el.aiModel.value.trim();
    return el.aiModelSelect.value;
  }

  function reasoningLevels() {
    const levels = catalog.reasoning[devIdFor(el.aiProvider.value)] || {};
    return levels[selectedModel()] || [];
  }

  function renderProviders(selectedId) {
    providerList = DualTranCatalog.buildProviderList(catalog.providers, catalog.modelsById);
    el.aiProvider.innerHTML = "";
    let dividerAdded = false;
    for (const provider of providerList) {
      if (!provider.builtIn && !dividerAdded) {
        const divider = document.createElement("option");
        divider.disabled = true;
        divider.value = MORE_VALUE;
        divider.textContent = "──── more ────";
        el.aiProvider.appendChild(divider);
        dividerAdded = true;
      }
      const option = document.createElement("option");
      option.value = provider.id;
      option.textContent = provider.name;
      el.aiProvider.appendChild(option);
    }
    const known = providerList.some((provider) => provider.id === selectedId);
    el.aiProvider.value = known ? selectedId : "openai";
  }

  function renderModels(selectedModelId) {
    const provider = providerById(el.aiProvider.value);
    const models = (provider && provider.models) || [];
    el.aiModelSelect.innerHTML = "";
    for (const model of models) {
      const option = document.createElement("option");
      option.value = model.id;
      option.textContent = model.label;
      el.aiModelSelect.appendChild(option);
    }
    const custom = document.createElement("option");
    custom.value = CUSTOM_VALUE;
    custom.textContent = "Custom…";
    el.aiModelSelect.appendChild(custom);

    const known = models.some((model) => model.id === selectedModelId);
    if (known) el.aiModelSelect.value = selectedModelId;
    else if (models.length && !selectedModelId) el.aiModelSelect.value = models[0].id;
    else el.aiModelSelect.value = CUSTOM_VALUE;
    if (!known && selectedModelId) el.aiModel.value = selectedModelId;
    toggleCustomRow();
  }

  function toggleCustomRow() {
    el.customModelRow.hidden = el.aiModelSelect.value !== CUSTOM_VALUE;
  }

  function renderReasoning(selected) {
    const levels = reasoningLevels();
    const values = DualTranCatalog.LEVEL_ORDER.filter((level) => levels.includes(level));
    el.aiReasoning.innerHTML = "";
    const entries = [["lowest", "Lowest (default)"], ["", "Provider default"]].concat(
      values.map((level) => [level, level])
    );
    for (const [value, label] of entries) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      el.aiReasoning.appendChild(option);
    }
    const known = entries.some(([value]) => value === selected);
    el.aiReasoning.value = known ? selected : "lowest";
  }

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
    renderProviders(s.aiProvider || "openai");
    renderModels(s.aiModel || "");
    el.aiEndpoint.value = s.aiEndpoint || "";
    renderReasoning(s.reasoningSetting || "lowest");
  }
  function current() {
    return {
      engine: el.engine.value,
      sourceLanguage: el.source.value.trim() || "auto",
      targetLanguage: el.target.value.trim() || "en",
      aiProvider: el.aiProvider.value,
      reasoningSetting: el.aiReasoning.value,
      reasoning: DualTranCatalog.resolveReasoning(el.aiReasoning.value, reasoningLevels()),
      aiEndpoint: el.aiEndpoint.value.trim(),
      aiModel: selectedModel(),
      aiKey: el.aiKey.value.trim(),
    };
  }
  function status(msg, isError, busy) {
    el.status.textContent = msg || "";
    el.status.classList.toggle("error", !!isError);
    el.status.classList.toggle("busy", !!busy);
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
    status("Translating…", false, true);
    el.result.textContent = "";
    el.copy.disabled = true;
    el.replace.disabled = true;
    el.insert.disabled = true;
    try {
      // translate.js speaks { endpoint, model, key }; settings use ai-prefixed names.
      const aiCfg = { endpoint: cfg.aiEndpoint, model: cfg.aiModel, key: cfg.aiKey, reasoning: cfg.reasoning };
      const out = cfg.engine === "ai"
        ? await DualTranTranslate.ai(text, cfg.sourceLanguage, cfg.targetLanguage, aiCfg)
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
    el.aiProvider.addEventListener("change", () => {
      const provider = providerById(el.aiProvider.value);
      if (provider) el.aiEndpoint.value = provider.api;
      renderModels("");
      renderReasoning(el.aiReasoning.value);
    });
    el.aiModelSelect.addEventListener("change", () => {
      toggleCustomRow();
      renderReasoning(el.aiReasoning.value);
    });
    el.aiModel.addEventListener("input", () => renderReasoning(el.aiReasoning.value));
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

  async function boot() {
    await loadCatalog();
    paintSettings();
    wire();
    if (typeof Office === "undefined") status("Preview mode — Office.js not available.");
  }

  if (typeof Office !== "undefined") {
    Office.onReady(() => { boot(); });
  } else {
    // plain-browser preview (no Office host): still usable for a quick check
    document.addEventListener("DOMContentLoaded", () => {
      boot();
    });
  }
})();

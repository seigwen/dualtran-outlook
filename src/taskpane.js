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

  // Interface language: Office's display language → browser language → English.
  // Office.context is only reliable inside onReady, so boot() re-resolves it.
  let i18n = window.DtI18n.init();
  let t = i18n.t;
  i18n.apply();

  /** Adopt the host's display language once Office is ready (no picker here). */
  function relocalize(preferred) {
    if (!preferred || window.DtI18n.detectLocale(preferred) === i18n.code) return;
    i18n = window.DtI18n.init(preferred);
    t = i18n.t;
    i18n.apply();
  }

  const SETTINGS_KEY = "dualtran.outlook.settings";
  const CATALOG_KEY = "dualtran.outlook.catalog";
  const CATALOG_TTL = 24 * 60 * 60 * 1000;
  const CUSTOM_VALUE = "__custom__";
  const DEFAULT_TARGET = window.DualTranLanguages.DEFAULT_TARGET;
  /** The unconfigured provider. A fresh install has no provider at all (SPEC 2.1.0). */
  const EMPTY_PROVIDER = "";
  /** Endpoint placeholder while nothing is configured. */
  const EMPTY_ENDPOINT_PLACEHOLDER = "https://";
  const REASONING_SENTINEL = "lowest";

  /** Projected models.dev data: { providers, modelsById, reasoning }. */
  let catalog = { providers: [], modelsById: {}, reasoning: {} };
  /** Built-ins and the catalogue in one list — what the dropdown shows. */
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

  /** True while the user has not chosen a provider: no request may leave the pane. */
  function providerMissing() {
    return el.aiProvider.value === EMPTY_PROVIDER;
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

  /**
   * The provider picker: a disabled, selected placeholder first (when nothing is
   * stored), then the one alphabetical list of real providers. The placeholder
   * is NOT part of that list and never sorts into it (SPEC 2.1.0 R2).
   */
  function renderProviders(selectedId) {
    providerList = DualTranCatalog.buildProviderList(catalog.providers, catalog.modelsById);
    const stored = typeof selectedId === "string" ? selectedId.trim() : "";
    const known = providerList.some((provider) => provider.id === stored);

    el.aiProvider.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = EMPTY_PROVIDER;
    placeholder.textContent = t("providerPlaceholder");
    placeholder.disabled = true;
    // Kept in the list only while it is the selection: once a provider is
    // chosen it disappears, so there is no way back to the empty state.
    if (!known) el.aiProvider.appendChild(placeholder);
    for (const provider of providerList) {
      const option = document.createElement("option");
      option.value = provider.id;
      option.textContent = provider.name;
      el.aiProvider.appendChild(option);
    }
    el.aiProvider.value = known ? stored : EMPTY_PROVIDER;
    placeholder.selected = !known;
    syncProviderState();
  }

  /**
   * Everything that depends on "is a provider chosen": the tooltip that explains
   * the disabled state, the model list, the endpoint placeholder and the
   * reasoning picker. One function so the two renderers cannot drift apart.
   */
  function syncProviderState() {
    const missing = providerMissing();
    el.aiProvider.title = missing ? t("providerRequiredError") : "";
    el.aiModelSelect.title = missing ? t("providerRequiredForModel") : "";
    el.aiModelSelect.disabled = missing;
    el.aiReasoning.disabled = missing;
    el.aiEndpoint.placeholder = missing ? EMPTY_ENDPOINT_PLACEHOLDER : "";
    if (missing) {
      // Unconfigured: no models, no stale endpoint — but the stored sentinel stays.
      el.aiModelSelect.innerHTML = "";
      el.customModelRow.hidden = true;
      el.aiEndpoint.value = "";
    }
  }

  /** The target picker: one option per language, codes as values, "Name (code)" as label. */
  function renderTargets(selectedCode) {
    const chosen = DualTranLanguages.normalizeTarget(selectedCode || DEFAULT_TARGET);
    el.target.innerHTML = "";
    for (const language of DualTranLanguages.TARGET_LANGUAGES) {
      const option = document.createElement("option");
      option.value = language.code;
      option.textContent = language.label;
      el.target.appendChild(option);
    }
    el.target.value = chosen;
  }

  function renderModels(selectedModelId) {
    const provider = providerById(el.aiProvider.value);
    if (!provider) { renderModelsForNoProvider(); return; }
    // Catalogue rows read alphabetically; the free-text entry is an action and
    // stays last (SPEC 2.1.0 R2).
    const models = ((provider && provider.models) || [])
      .slice()
      .sort(
        (a, b) =>
          a.label.localeCompare(b.label, "en", { numeric: true, sensitivity: "base" }) ||
          a.id.localeCompare(b.id)
      );
    el.aiModelSelect.innerHTML = "";
    for (const model of models) {
      const option = document.createElement("option");
      option.value = model.id;
      option.textContent = model.label;
      el.aiModelSelect.appendChild(option);
    }
    const custom = document.createElement("option");
    custom.value = CUSTOM_VALUE;
    custom.textContent = t("customModel");
    el.aiModelSelect.appendChild(custom);

    const known = models.some((model) => model.id === selectedModelId);
    if (known) el.aiModelSelect.value = selectedModelId;
    else if (models.length && !selectedModelId) el.aiModelSelect.value = sortedFirstModel(models);
    else el.aiModelSelect.value = CUSTOM_VALUE;
    if (!known && selectedModelId) el.aiModel.value = selectedModelId;
    toggleCustomRow();
  }

  /** No provider: the model picker is empty, disabled and says what is missing. */
  function renderModelsForNoProvider() {
    el.aiModelSelect.innerHTML = "";
    el.customModelRow.hidden = true;
    syncProviderState();
  }

  /** The model the picker shows when nothing is stored: the first catalogue row it displays. */
  function sortedFirstModel(models) {
    return models
      .slice()
      .sort((a, b) => a.label.localeCompare(b.label, "en", { numeric: true, sensitivity: "base" }) || a.id.localeCompare(b.id))[0].id;
  }

  function toggleCustomRow() {
    el.customModelRow.hidden = el.aiModelSelect.value !== CUSTOM_VALUE;
  }

  function renderReasoning(selected) {
    const levels = reasoningLevels();
    const values = DualTranCatalog.LEVEL_ORDER.filter((level) => levels.includes(level));
    el.aiReasoning.innerHTML = "";
    const entries = [[REASONING_SENTINEL, t("reasoningLowest")], ["", t("reasoningDefault")]].concat(
      values.map((level) => [level, level])
    );
    for (const [value, label] of entries) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      el.aiReasoning.appendChild(option);
    }
    const known = entries.some(([value]) => value === selected);
    el.aiReasoning.value = known ? selected : REASONING_SENTINEL;
    el.aiReasoning.disabled = providerMissing();
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
    renderTargets(s.targetLanguage || DEFAULT_TARGET);
    // A stored provider is painted as stored; a fresh install stays empty.
    const stored = typeof s.aiProvider === "string" ? s.aiProvider : EMPTY_PROVIDER;
    renderProviders(stored);
    if (providerMissing()) {
      // Nothing configured: keep the sentinel, paint no endpoint, no models.
      renderReasoning(s.reasoningSetting || REASONING_SENTINEL);
      return;
    }
    renderModels(s.aiModel || "");
    el.aiEndpoint.value = s.aiEndpoint || "";
    renderReasoning(s.reasoningSetting || REASONING_SENTINEL);
  }
  function current() {
    return {
      engine: el.engine.value,
      sourceLanguage: el.source.value.trim() || "auto",
      targetLanguage: DualTranLanguages.normalizeTarget(el.target.value),
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
    if (!item) { status(t("statusOpenMessage"), true); return; }
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
    if (!text) { status(t("statusNothing"), true); return; }
    // Nothing configured: say so and leave before any network call (SPEC 2.1.0).
    if (providerMissing()) { status(t("providerRequiredError"), true); return; }
    const cfg = current();
    saveSettings(cfg);
    el.translate.disabled = true;
    status(t("statusTranslating"), false, true);
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
      status(out ? t("statusDone") : t("statusEmpty"));
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
      const cb = (r) => {{
        const ok = r.status === Office.AsyncResultStatus.Succeeded;
        status(ok ? t("statusWritten") : t("statusWriteFailed"), !ok);
      }};
      if (mode === "replace") item.body.setAsync(out, { coercionType: Office.CoercionType.Text }, cb);
      else item.body.setSelectedDataAsync(out, { coercionType: Office.CoercionType.Text }, cb);
    });
  }

  function wire() {
    el.aiProvider.addEventListener("change", () => {
      // The placeholder is gone once a provider is chosen, so this only ever
      // sees real providers; an empty value would mean "stay put".
      if (providerMissing()) return;
      const provider = providerById(el.aiProvider.value);
      el.aiEndpoint.value = DualTranCatalog.resolveEndpointOnProviderChange(
        el.aiEndpoint.value,
        provider ? provider.api : ""
      );
      syncProviderState();
      renderModels("");
      renderReasoning(el.aiReasoning.value);
      updateActionLabels();
    });
    el.aiModelSelect.addEventListener("change", () => {
      toggleCustomRow();
      renderReasoning(el.aiReasoning.value);
    });
    el.aiModel.addEventListener("input", () => renderReasoning(el.aiReasoning.value));
    el.translate.addEventListener("click", translateNow);
    el.copy.addEventListener("click", () => navigator.clipboard.writeText(el.result.textContent).then(() => status(t("statusCopied"))));
    el.load.addEventListener("click", async () => {
      const selected = await getSelected();
      if (selected) { el.input.value = selected; status(t("statusLoadedSelection")); return; }
      try { el.input.value = await getBody(); status(t("statusLoadedAll")); }
      catch (e) { status(String(e.message || e), true); }
    });
    el.replace.addEventListener("click", () => writeBody("replace"));
    el.insert.addEventListener("click", () => writeBody("insert"));
    el.settings.addEventListener("click", () => { el.panel.hidden = !el.panel.hidden; });
    el.save.addEventListener("click", () => { saveSettings(current()); status(t("statusSaved")); });
    el.input.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) translateNow();
    });
  }

  /**
   * The Translate button explains, while nothing is configured, why it would do
   * nothing: same string as the status it would show (SPEC 2.1.0). It stays
   * clickable so the click itself can say so.
   */
  function updateActionLabels() {
    el.translate.title = providerMissing() ? t("providerRequiredError") : "";
  }

  async function boot() {
    relocalize(typeof Office !== "undefined" && Office.context ? Office.context.displayLanguage : "");
    await loadCatalog();
    paintSettings();
    updateActionLabels();
    wire();
    if (typeof Office === "undefined") status(t("statusPreview"));
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

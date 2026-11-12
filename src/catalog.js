/**
 * catalog.js — provider list for hosts that only speak OpenAI-compatible HTTP.
 *
 * Two layers, the same split every DualTran surface uses:
 *   • BUILTINS  — curated ids, names and endpoints, always first in the list;
 *   • models.dev — the long tail (200+ providers), fetched once and cached.
 *
 * models.dev's own entries for the big names carry no `api` (the SDK infers it),
 * so a built-in still pulls its *models* from the catalogue entry it aliases —
 * that is where the reasoning declarations live.
 *
 * Identical to the copy in dualtran-vscode/src/catalog.js; keep the two in sync.
 * Exposed as `window.DualTranCatalog` and as `module.exports` so Node tests can
 * require it without a DOM.
 */
(function (root, factory) {
  "use strict";
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.DualTranCatalog = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const MODELS_DEV_URL = "https://models.dev/api.json";

  /** Curated providers: id → endpoint. `dev` is the models.dev id it aliases. */
  const BUILTINS = [
    { id: "openai", name: "OpenAI", dev: "openai", api: "https://api.openai.com/v1", models: ["gpt-4o-mini"] },
    { id: "google-gemini", name: "Google Gemini", dev: "google", api: "https://generativelanguage.googleapis.com/v1beta/openai", models: ["gemini-2.5-flash"] },
    { id: "deepseek", name: "DeepSeek", dev: "deepseek", api: "https://api.deepseek.com/v1", models: ["deepseek-chat"] },
    { id: "moonshot", name: "Moonshot", dev: "moonshotai", api: "https://api.moonshot.ai/v1", models: ["kimi-k2-0905-preview"] },
    { id: "zhipu", name: "Zhipu", dev: "zhipuai", api: "https://open.bigmodel.cn/api/paas/v4", models: ["glm-4.6"] },
    { id: "groq", name: "Groq", dev: "groq", api: "https://api.groq.com/openai/v1", models: ["llama-3.3-70b-versatile"] },
    { id: "mistral", name: "Mistral", dev: "mistral", api: "https://api.mistral.ai/v1", models: ["mistral-large-latest"] },
    { id: "xai", name: "xAI", dev: "grok", api: "https://api.x.ai/v1", models: ["grok-4"] },
    { id: "openrouter", name: "OpenRouter", dev: "openrouter", api: "https://openrouter.ai/api/v1", models: ["openai/gpt-4o-mini"] },
    { id: "siliconflow", name: "SiliconFlow", dev: "siliconflow", api: "https://api.siliconflow.cn/v1", models: ["deepseek-ai/DeepSeek-V3"] },
    { id: "ollama", name: "Ollama (local)", dev: "ollama", api: "http://localhost:11434/v1", models: ["llama3.2"] },
  ];

  const LEVEL_ORDER = ["none", "minimal", "low", "medium", "high", "xhigh", "max"];

  /**
   * Curated endpoints for models.dev entries that carry no `api` of their own.
   * models.dev omits it for the big names because their official SDKs infer a
   * base URL; a BYOK pane that talks plain HTTP has to know it. Without an
   * entry here such a provider cannot be reached and stays out of the picker
   * (see buildProviderList) — a pickable provider that fills nothing would
   * break SPEC R3. Never invent a URL: only well-known, documented bases.
   */
  const MODELS_DEV_ENDPOINTS = {
    anthropic: "https://api.anthropic.com/v1",
  };

  /** Provider labels sort the way the picker reads them: case- and accent-blind, numbers in order. */
  const collate = new Intl.Collator("en", { numeric: true, sensitivity: "base" });

  /** Effort values a model declares, in a stable order; a toggle model reports ["high"]. */
  function levelsOf(model) {
    const options = Array.isArray(model && model.reasoning_options) ? model.reasoning_options : [];
    const effort = options.find((option) => option && option.type === "effort");
    if (effort && Array.isArray(effort.values)) {
      const values = effort.values.filter((value) => typeof value === "string");
      const ordered = LEVEL_ORDER.filter((level) => values.includes(level));
      return ordered.concat(values.filter((value) => !LEVEL_ORDER.includes(value)));
    }
    return options.some((option) => option && option.type === "toggle") ? ["high"] : [];
  }

  /** Trim models.dev's api.json to what a dropdown needs. */
  function projectCatalog(raw) {
    const providers = [];
    const modelsById = {};
    const reasoning = {};
    for (const [id, entry] of Object.entries(raw || {})) {
      if (!entry || typeof entry !== "object") continue;
      const models = [];
      const levels = {};
      for (const [modelId, model] of Object.entries(entry.models || {})) {
        models.push({ id: modelId, label: (model && model.name) || modelId });
        levels[modelId] = levelsOf(model);
      }
      reasoning[id] = levels;
      modelsById[id] = models;
      if (entry.api && models.length) {
        providers.push({ id, name: entry.name || id, api: entry.api, models });
      }
    }
    providers.sort((a, b) => a.name.localeCompare(b.name));
    return { providers, modelsById, reasoning };
  }

  /** providerList entry → the models.dev shape projectCatalog reads ({ models: { id: { name } } }). */
  function projectEntries(lists) {
    const raw = {};
    for (const entry of lists || []) {
      const models = {};
      for (const model of entry.models || []) models[model.id] = { id: model.id, name: model.label };
      raw[entry.id] = { name: entry.name, api: entry.api, models };
    }
    return projectCatalog(raw);
  }

  /**
   * Built-ins and the catalogue in ONE list, alphabetical by visible label; alias
   * twins removed. A models.dev entry with no known endpoint is left out: the
   * picker only offers providers whose endpoint it can actually fill (SPEC R3).
   */
  function buildProviderList(catalogue, modelsById) {
    const byId = new Map((catalogue || []).map((provider) => [provider.id, provider]));
    const fallbackModels = modelsById || {};
    const claimed = new Set(BUILTINS.map((entry) => entry.dev));
    const out = BUILTINS.map((entry) => {
      const dev = byId.get(entry.dev);
      const models = (dev && dev.models) || fallbackModels[entry.dev] || (entry.models || []).map((id) => ({ id, label: id }));
      return { id: entry.id, name: entry.name, api: entry.api, models, builtIn: true };
    });
    for (const provider of catalogue || []) {
      if (claimed.has(provider.id)) continue;
      if (BUILTINS.some((entry) => entry.id === provider.id)) continue;
      const api = provider.api || MODELS_DEV_ENDPOINTS[provider.id] || "";
      if (!api) continue;
      out.push({ id: provider.id, name: provider.name, api, models: provider.models, builtIn: false });
    }
    // One alphabetical list: a built-in and a catalogue entry sit side by side,
    // so a user looking for "Groq" does not have to know which half it is in.
    // Ties (same label from two ids) fall back to the id for a stable order.
    return out.sort((a, b) => collate.compare(a.name, b.name) || collate.compare(a.id, b.id));
  }

  /**
   * Endpoint after a provider change: an empty (or whitespace-only) endpoint is
   * filled with the provider's API; anything the user typed wins, so a proxy or
   * gateway URL is never clobbered. A provider without an api fills nothing.
   */
  function resolveEndpointOnProviderChange(currentEndpoint, providerApi) {
    const current = typeof currentEndpoint === "string" ? currentEndpoint.trim() : "";
    if (current) return current;
    const api = typeof providerApi === "string" ? providerApi.trim() : "";
    return api;
  }

  /** Turn the stored setting into the value to send; nothing declared → no parameter. */
  function resolveReasoning(setting, levels) {
    if (setting === "default") return "";
    if (setting !== "lowest") return setting || "";
    const values = Array.isArray(levels) ? levels : [];
    for (const level of ["none", "minimal", "low"]) {
      if (values.includes(level)) return level;
    }
    return "";
  }

  return { MODELS_DEV_URL, BUILTINS, MODELS_DEV_ENDPOINTS, LEVEL_ORDER, levelsOf, projectCatalog, projectEntries, buildProviderList, resolveEndpointOnProviderChange, resolveReasoning };
});

/**
 * translate.js — the two engines, framework-free so the taskpane has no build step.
 *
 * google(): the public gtx endpoint; it answers with CORS headers, so a browser
 * page can call it directly (verified). Batches long text into <2k-char chunks.
 * ai(): any OpenAI-compatible /chat/completions endpoint with the user's own key.
 */
(function (global) {
  "use strict";

  const GTX = "https://translate.googleapis.com/translate_a/single";
  const MAX_CHUNK = 1800;

  function chunk(text, size) {
    const out = [];
    let buf = "";
    for (const part of text.split(/(?<=[.!?。！？\n])/)) {
      if ((buf + part).length > size && buf) { out.push(buf); buf = part; }
      else buf += part;
    }
    if (buf) out.push(buf);
    return out.length ? out : [text];
  }

  async function google(text, sourceLang, targetLang) {
    const chunks = chunk(text, MAX_CHUNK);
    const parts = [];
    for (const c of chunks) {
      const url = `${GTX}?client=gtx&sl=${encodeURIComponent(sourceLang || "auto")}` +
                  `&tl=${encodeURIComponent(targetLang)}&dt=t&q=${encodeURIComponent(c)}`;
      const res = await fetch(url);
      if (!res.ok) throw new Error("Google translation failed: HTTP " + res.status);
      const data = await res.json();
      parts.push((data[0] || []).map((seg) => seg[0]).join(""));
    }
    return parts.join("\n");
  }

  async function ai(text, sourceLang, targetLang, cfg) {
    if (!cfg || !cfg.endpoint || !cfg.model || !cfg.key) {
      throw new Error("AI mode needs an endpoint, a model and an API key (⚙).");
    }
    const base = cfg.endpoint.replace(/\/+$/, "");
    const url = base + "/chat/completions";
    const body = {
      model: cfg.model,
      temperature: 0.2,
      messages: [
        { role: "system", content: `Translate into ${targetLang}. Output only the translation, keep line breaks.` },
        { role: "user", content: text },
      ],
    };
    if (cfg.reasoning) body.reasoning_effort = cfg.reasoning;
    const post = (payload) =>
      fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: "Bearer " + cfg.key },
        body: JSON.stringify(payload),
      });

    let res = await post(body);
    if (!res.ok && res.status === 400 && body.reasoning_effort) {
      // A gateway that does not know the parameter gets the request again without it.
      const detail = await res.text().catch(() => "");
      if (/reasoning|effort|unsupported|unrecognized/i.test(detail)) {
        delete body.reasoning_effort;
        res = await post(body);
        if (res.ok) {
          const retried = await res.json();
          return (retried.choices && retried.choices[0] && retried.choices[0].message.content || "").trim();
        }
      }
      throw new Error("AI translation failed: HTTP " + res.status + (detail ? " — " + detail.slice(0, 160) : ""));
    }
    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      throw new Error("AI translation failed: HTTP " + res.status + (detail ? " — " + detail.slice(0, 160) : ""));
    }
    const data = await res.json();
    return (data.choices && data.choices[0] && data.choices[0].message.content || "").trim();
  }

  global.DualTranTranslate = { google, ai, chunk };
})(window);

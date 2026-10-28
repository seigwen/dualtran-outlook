/**
 * run.mjs — Outlook task-pane E2E (host API stubbed, real DOM, real fetch).
 *
 * Six scenarios, the same contract every DualTran host follows:
 *   1. settings shape   — provider select with built-ins → "──── more ────" → catalogue,
 *                         reasoning starts with Lowest + Default
 *   2. Google translate — selection → translated text in #result
 *   3. AI request body  — endpoint/model/key/reasoning reach the wire intact
 *   4. loading visible  — spinner class while the request is in flight, gone after
 *   5. error path       — 401 shows a readable error, spinner cleared, button enabled
 *   6. persistence      — reload keeps provider/model/reasoning
 *
 * Artifacts (screenshots, console, requests) land in e2e-artifacts/<scenario>/.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import { startMockServer } from "./mock-server.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ARTIFACTS = path.join(HERE, "..", "e2e-artifacts");
const FIXTURE = JSON.parse(fs.readFileSync(path.join(HERE, "fixtures", "modelsdev.json"), "utf8"));

const OFFICE_STUB = (displayLanguage = "en-US") => `
  window.Office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Text: "text" },
    context: {
      displayLanguage: ${JSON.stringify(displayLanguage)},
      mailbox: {
        item: {
          body: {
            getAsync(_type, cb) {
              window.__bodyReads = (window.__bodyReads || 0) + 1;
              setTimeout(() => cb({ status: "succeeded", value: "Hello from the message" }), 5);
            },
            setAsync(value, _opts, cb) {
              window.__writtenBody = value;
              setTimeout(() => cb({ status: "succeeded" }), 5);
            },
            setSelectedDataAsync(value, _opts, cb) {
              window.__insertedText = value;
              setTimeout(() => cb({ status: "succeeded" }), 5);
            },
          },
          getSelectedDataAsync(_type, cb) {
            setTimeout(() => cb({ status: "succeeded", value: { data: "Selected text" } }), 5);
          },
        },
      },
    },
    onReady(callback) {
      setTimeout(callback, 0);
    },
  };
`;

const scenarios = [];
const scenario = (name, run) => scenarios.push({ name, run });

async function openPage(browser, origin, options = {}) {
  const context = await browser.newContext({ viewport: { width: 520, height: 900 }, ...options });
  const consoleLines = [];
  const page = await context.newPage();
  page.on("console", (message) => consoleLines.push(`${message.type()}: ${message.text()}`));
  page.on("pageerror", (error) => consoleLines.push(`pageerror: ${error.message}`));

  // Office.context.displayLanguage is what the pane actually adopts.
  await page.addInitScript(OFFICE_STUB(options.displayLanguage));
  await page.route("https://appsforoffice.microsoft.com/**", (route) =>
    route.fulfill({ status: 200, contentType: "text/javascript", body: "/* office.js stubbed */" })
  );
  await page.route("https://models.dev/api.json", (route) =>
    route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(FIXTURE) })
  );
  await page.route("**/translate_a/single**", async (route) => {
    const url = new URL(route.request().url());
    await route.fulfill({
      status: 200,
      contentType: "application/json",
      body: JSON.stringify([[["译文", url.searchParams.get("q") || "", null, null, 10]], null, "en"]),
    });
  });

  await page.goto(`${origin}/taskpane.html`);
  // <option> inside a closed <select> is never "visible"; wait on the DOM instead.
  await page.waitForFunction(() => document.querySelectorAll("#aiProvider option").length > 0);
  // Catalogue fetch is async; wait until the divider exists.
  await page.waitForFunction(() =>
    Array.from(document.querySelectorAll("#aiProvider option")).some((option) => option.disabled)
  );
  return { page, context, consoleLines };
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

scenario("settings-shape", async ({ browser, server }) => {
  const { page, context } = await openPage(browser, server.origin);
  const shape = await page.evaluate(() => {
    const provider = document.getElementById("aiProvider");
    const options = Array.from(provider.options);
    const divider = options.findIndex((option) => option.disabled);
    const reasoning = document.getElementById("aiReasoning");
    return {
      providers: options.length,
      firstThree: options.slice(0, 3).map((option) => option.value),
      dividerText: divider >= 0 ? options[divider].textContent.trim() : null,
      afterDivider: divider >= 0 ? options.slice(divider + 1).map((option) => option.value) : [],
      reasoning: Array.from(reasoning.options).map((option) => option.value),
      reasoningDefault: reasoning.value,
    };
  });
  assert(shape.providers > 11, `expected built-ins + catalogue, got ${shape.providers} options`);
  assert(shape.firstThree[0] === "openai", `built-ins must lead, got ${shape.firstThree.join(",")}`);
  assert(
    shape.dividerText && shape.dividerText.includes("more"),
    `divider must read "──── more ────", got ${JSON.stringify(shape.dividerText)}`
  );
  assert(shape.afterDivider.includes("orcarouter"), "catalogue providers must follow the divider");
  assert(shape.reasoning[0] === "lowest", `reasoning must default to lowest, got ${shape.reasoning[0]}`);
  assert(shape.reasoningDefault === "lowest", `selected value must be lowest, got ${shape.reasoningDefault}`);
  await page.screenshot({ path: path.join(ARTIFACTS, "settings-shape", "screenshot.png") });
  await context.close();
});

scenario("google-translate", async ({ browser, server }) => {
  const { page, context } = await openPage(browser, server.origin);
  await page.fill("#input", "Hello world");
  await page.selectOption("#engine", "google");
  await page.click("#translate");
  await page.waitForFunction(() => document.getElementById("result").textContent.trim().length > 0, null, {
    timeout: 5000,
  });
  const result = await page.textContent("#result");
  const status = await page.textContent("#status");
  assert(result.includes("译文"), `expected the translated text, got ${JSON.stringify(result)}`);
  assert(status.includes("Done"), `expected a Done status, got ${JSON.stringify(status)}`);
  await page.screenshot({ path: path.join(ARTIFACTS, "google-translate", "screenshot.png") });
  await context.close();
});

scenario("ai-request-body", async ({ browser, server }) => {
  const { page, context } = await openPage(browser, server.origin);
  await page.click("#settings");
  await page.selectOption("#aiProvider", "openai");
  await page.selectOption("#aiModelSelect", "gpt-5");
  await page.fill("#aiEndpoint", `${server.origin}/v1`);
  await page.selectOption("#aiReasoning", "low");
  await page.fill("#aiKey", "test-key");
  await page.selectOption("#engine", "ai");
  await page.fill("#input", "Hello world");
  server.clear();
  await page.click("#translate");
  await page.waitForFunction(() => document.getElementById("result").textContent.trim().length > 0, null, {
    timeout: 5000,
  });
  const call = server.requests.find((entry) => entry.url === "/v1/chat/completions");
  assert(call, "no AI request reached the mock endpoint");
  assert(call.body.model === "gpt-5", `wrong model: ${call.body.model}`);
  assert(call.body.reasoning_effort === "low", `wrong reasoning_effort: ${call.body.reasoning_effort}`);
  assert(
    call.headers.authorization === "Bearer test-key",
    `wrong authorization header: ${call.headers.authorization}`
  );
  assert(
    call.body.messages.some((message) => /Translate into zh-CN/.test(message.content)),
    "prompt lost the target language"
  );
  fs.writeFileSync(path.join(ARTIFACTS, "ai-request-body", "requests.json"), JSON.stringify(server.requests, null, 2));
  await page.screenshot({ path: path.join(ARTIFACTS, "ai-request-body", "screenshot.png") });
  await context.close();
});

scenario("loading-visible", async ({ browser, server }) => {
  const { page, context } = await openPage(browser, server.origin);
  await page.click("#settings");
  await page.selectOption("#aiProvider", "openai");
  await page.selectOption("#aiModelSelect", "gpt-5");
  await page.fill("#aiEndpoint", `${server.origin}/v1`);
  await page.fill("#aiKey", "test-key");
  await page.selectOption("#engine", "ai");
  await page.fill("#input", "Hello world");
  server.setAi({ delayMs: 1200 });
  await page.click("#translate");
  await page.waitForFunction(() => document.getElementById("status").classList.contains("busy"), null, {
    timeout: 1500,
  });
  await page.screenshot({ path: path.join(ARTIFACTS, "loading-visible", "screenshot-while-busy.png") });
  await page.waitForFunction(() => !document.getElementById("status").classList.contains("busy"), null, {
    timeout: 6000,
  });
  const after = await page.evaluate(() => ({
    busy: document.getElementById("status").classList.contains("busy"),
    result: document.getElementById("result").textContent.trim(),
  }));
  assert(!after.busy, "spinner stayed after the request finished");
  assert(after.result.length > 0, "no result after the delayed request");
  await context.close();
});

scenario("error-path", async ({ browser, server }) => {
  const { page, context } = await openPage(browser, server.origin);
  await page.click("#settings");
  await page.selectOption("#aiProvider", "openai");
  await page.selectOption("#aiModelSelect", "gpt-5");
  await page.fill("#aiEndpoint", `${server.origin}/v1`);
  await page.fill("#aiKey", "bad-key");
  await page.selectOption("#engine", "ai");
  await page.fill("#input", "Hello world");
  server.setAi({ status: 401, delayMs: 0 });
  await page.click("#translate");
  await page.waitForFunction(
    () => document.getElementById("status").classList.contains("error"),
    null,
    { timeout: 5000 }
  );
  const state = await page.evaluate(() => ({
    text: document.getElementById("status").textContent,
    busy: document.getElementById("status").classList.contains("busy"),
    disabled: document.getElementById("translate").disabled,
    result: document.getElementById("result").textContent,
  }));
  assert(/401/.test(state.text), `error status must mention 401, got ${JSON.stringify(state.text)}`);
  assert(!state.busy, "spinner must be cleared on failure");
  assert(!state.disabled, "translate button must be usable again");
  assert(state.result === "", "a failed request must not leave a result behind");
  await page.screenshot({ path: path.join(ARTIFACTS, "error-path", "screenshot.png") });
  await context.close();
});

scenario("persistence", async ({ browser, server }) => {
  const { page, context } = await openPage(browser, server.origin);
  await page.click("#settings");
  await page.selectOption("#aiProvider", "orcarouter");
  await page.selectOption("#aiModelSelect", "orcarouter/fusion-mini");
  await page.selectOption("#aiReasoning", "high");
  await page.fill("#aiEndpoint", `${server.origin}/v1`);
  await page.fill("#target", "ja");
  await page.click("#saveSettings");
  await page.reload();
  await page.waitForFunction(() => document.querySelectorAll("#aiProvider option").length > 0);
  await page.waitForFunction(() =>
    Array.from(document.querySelectorAll("#aiProvider option")).some((option) => option.disabled)
  );
  const restored = await page.evaluate(() => ({
    provider: document.getElementById("aiProvider").value,
    model: document.getElementById("aiModelSelect").value,
    reasoning: document.getElementById("aiReasoning").value,
    target: document.getElementById("target").value,
    endpoint: document.getElementById("aiEndpoint").value,
  }));
  assert(restored.provider === "orcarouter", `provider not persisted: ${restored.provider}`);
  assert(restored.model === "orcarouter/fusion-mini", `model not persisted: ${restored.model}`);
  assert(restored.reasoning === "high", `reasoning not persisted: ${restored.reasoning}`);
  assert(restored.target === "ja", `target language not persisted: ${restored.target}`);
  await page.screenshot({ path: path.join(ARTIFACTS, "persistence", "screenshot.png") });
  await context.close();
});

scenario("i18n-locale", async ({ browser, server }) => {
  const read = (code) => JSON.parse(fs.readFileSync(path.join(HERE, "..", "locales", code, "messages.json"), "utf8"));
  const de = read("de");

  // Office's display language is adopted even though the page renders before onReady.
  const { page, context } = await openPage(browser, server.origin, { displayLanguage: "de-DE" });
  const ui = await page.evaluate(() => ({
    translate: document.getElementById("translate").textContent,
    copy: document.getElementById("copy").textContent,
    load: document.getElementById("load").textContent,
    engineGoogle: document.querySelector('#engine option[value="google"]').textContent,
    engineAi: document.querySelector('#engine option[value="ai"]').textContent,
    placeholder: document.getElementById("input").getAttribute("placeholder"),
    providerLabel: document.querySelector('[data-i18n="settingProvider"]').textContent,
    hint: document.querySelector('[data-i18n="settingsHint"]').textContent,
    reasoningFirst: document.getElementById("aiReasoning").options[0].textContent,
    divider: Array.from(document.getElementById("aiProvider").options).find((option) => option.disabled)?.textContent,
    dictDir: document.documentElement.getAttribute("dir"),
    htmlLang: document.documentElement.getAttribute("lang"),
  }));
  assert(ui.translate === de.btnTranslate.message, `translate must be German, got "${ui.translate}"`);
  assert(ui.copy === de.btnCopy.message, `copy must be German, got "${ui.copy}"`);
  assert(ui.load === de.btnLoad.message, `load must be German, got "${ui.load}"`);
  assert(ui.engineGoogle === de.engineGoogle.message, `engine option must be German, got "${ui.engineGoogle}"`);
  assert(ui.engineAi === de.engineAi.message, `AI option must be German, got "${ui.engineAi}"`);
  assert(ui.placeholder === de.inputPlaceholder.message, `placeholder must be German, got "${ui.placeholder}"`);
  assert(ui.providerLabel === de.settingProvider.message, `settings label must be German, got "${ui.providerLabel}"`);
  assert(ui.hint === de.settingsHint.message, `hint must be German, got "${ui.hint}"`);
  assert(ui.reasoningFirst === de.reasoningLowest.message, `reasoning label must be German, got "${ui.reasoningFirst}"`);
  assert(ui.divider === de.providerDivider.message, `catalogue divider must be German, got "${ui.divider}"`);
  assert(ui.dictDir === null, `German must not set dir, got "${ui.dictDir}"`);
  assert(ui.htmlLang === "de", `<html lang> must follow the locale, got "${ui.htmlLang}"`);

  // No label may render its raw key or stay empty.
  const leaks = await page.evaluate(() =>
    Array.from(document.querySelectorAll("[data-i18n]"))
      .filter((node) => {
        const key = node.getAttribute("data-i18n");
        const text = (node.textContent || "").trim();
        return !text || text === key;
      })
      .map((node) => node.getAttribute("data-i18n"))
  );
  assert(leaks.length === 0, `untranslated labels: ${leaks.join(", ")}`);

  // Dynamic status text is translated too.
  await page.click("#translate");
  await page.waitForFunction(() => document.getElementById("status").textContent.trim().length > 0);
  const statusText = await page.evaluate(() => document.getElementById("status").textContent);
  assert(statusText === de.statusNothing.message, `status must be German, got "${statusText}"`);
  await page.screenshot({ path: path.join(ARTIFACTS, "i18n-locale", "de.png") });
  await context.close();

  // Arabic flips the pane and mirrors the arrow.
  const rtl = await openPage(browser, server.origin, { displayLanguage: "ar-SA" });
  const ar = read("ar");
  const flipped = await rtl.page.evaluate(() => ({
    dir: document.documentElement.getAttribute("dir"),
    arrow: getComputedStyle(document.querySelector(".arrow")).transform,
    button: document.getElementById("translate").textContent,
  }));
  assert(flipped.dir === "rtl", `Arabic must set dir=rtl, got "${flipped.dir}"`);
  assert(flipped.arrow.includes("-1"), `arrow must be mirrored, got "${flipped.arrow}"`);
  assert(flipped.button === ar.btnTranslate.message, `Arabic button expected, got "${flipped.button}"`);
  await rtl.page.screenshot({ path: path.join(ARTIFACTS, "i18n-locale", "ar.png") });
  await rtl.context.close();

  // A language we do not ship falls back to English, never to keys.
  const en = read("en");
  const fallback = await openPage(browser, server.origin, { displayLanguage: "xx-YY" });
  const fallbackText = await fallback.page.evaluate(() => document.getElementById("translate").textContent);
  assert(fallbackText === en.btnTranslate.message, `unknown locale must fall back to English, got "${fallbackText}"`);
  await fallback.context.close();
});

async function main() {
  fs.rmSync(ARTIFACTS, { recursive: true, force: true });
  for (const entry of scenarios) fs.mkdirSync(path.join(ARTIFACTS, entry.name), { recursive: true });

  const server = await startMockServer();
  const browser = await chromium.launch({ headless: true });
  const failures = [];

  for (const entry of scenarios) {
    const started = Date.now();
    process.stdout.write(`▶ ${entry.name} … `);
    try {
      await entry.run({ browser, server });
      console.log(`ok (${Date.now() - started} ms)`);
    } catch (error) {
      failures.push({ name: entry.name, error: error && error.message ? error.message : String(error) });
      console.log(`FAILED\n  ${error && error.stack ? error.stack.split("\n").slice(0, 4).join("\n  ") : error}`);
      fs.writeFileSync(path.join(ARTIFACTS, entry.name, "error.txt"), String(error && error.stack ? error.stack : error));
    } finally {
      server.clear();
    }
  }

  await browser.close();
  await server.close();

  if (failures.length) {
    console.error(`\n${failures.length}/${scenarios.length} scenarios failed`);
    process.exit(1);
  }
  console.log(`\n${scenarios.length}/${scenarios.length} scenarios passed`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

#!/usr/bin/env node
/**
 * fill-provider-keys.mjs — top up the three SPEC 2.1.0 provider keys in every
 * locale, keeping every existing translation untouched.
 *
 * This is the repo's `scripts/locale-kit.mjs --fill-missing` mechanism, run for
 * one key-set in this session's environment: locale-kit's two sources are the
 * browser extension's catalogue (../DualTran-extension/src/_locales — not
 * present in this workspace) and a local OpenAI-compatible model
 * (http://127.0.0.1:11434/v1 — not reachable here). So it borrows locale-kit's
 * "append only, never rewrite" contract and its --check validation, and sources
 * the strings from:
 *   1. the extension translation memory, when that directory exists;
 *   2. a local OpenAI-compatible model, when OLLAMA_BASE_URL answers;
 *   3. otherwise the hand-written table below (human-authored for this change),
 *      which is what ends up in the catalogues here.
 *
 * zh_CN / zh_TW are hand-written in the repo and already filled — never touched.
 *
 *   node scripts/fill-provider-keys.mjs            # fill missing keys
 *   node scripts/locale-kit.mjs --check            # validate the result
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const LOCALES = path.join(ROOT, "locales");
const EXT = process.env.EXT_LOCALES || path.resolve(ROOT, "..", "..", "DualTran-extension", "src", "_locales");
const BASE = (process.env.OLLAMA_BASE_URL || "http://127.0.0.1:11434/v1").replace(/\/+$/, "");
const MODEL = process.env.OLLAMA_MODEL || "deepseek-v4.1-flash:cloud";

const EN = {
  providerPlaceholder: "Select a provider",
  providerRequiredForModel: "Select a provider first",
  providerRequiredError: "Choose an AI provider in settings before translating.",
};
const HANDWRITTEN = new Set(["en", "zh_CN", "zh_TW"]);

/**
 * Human-authored translations for this change. Every line follows the tone of
 * the catalogue that already ships: settingProvider + settingApiKey wording,
 * the locale's own "first/before" idiom, no trailing period added or removed
 * relative to the source.
 */
const TABLE = {
  af: ["Kies ’n verskaffer", "Kies eers ’n verskaffer", "Kies ’n KI-verskaffer in instellings voordat jy vertaal."],
  ar: ["اختر مزودًا", "اختر مزودًا أولًا", "اختر مزود ذكاء اصطناعي في الإعدادات قبل الترجمة."],
  ca: ["Selecciona un proveïdor", "Selecciona primer un proveïdor", "Tria un proveïdor d'IA als ajustos abans de traduir."],
  cs: ["Vyberte poskytovatele", "Nejprve vyberte poskytovatele", "Před překladem vyberte v nastavení poskytovatele AI."],
  da: ["Vælg en udbyder", "Vælg en udbyder først", "Vælg en AI-udbyder i indstillingerne, før du oversætter."],
  de: ["Anbieter auswählen", "Zuerst einen Anbieter auswählen", "Wähle vor dem Übersetzen einen KI-Anbieter in den Einstellungen."],
  el: ["Επιλέξτε πάροχο", "Επιλέξτε πρώτα πάροχο", "Επιλέξτε πάροχο AI στις ρυθμίσεις πριν μεταφράσετε."],
  es: ["Selecciona un proveedor", "Selecciona primero un proveedor", "Elige un proveedor de IA en los ajustes antes de traducir."],
  fa: ["یک ارائه‌دهنده انتخاب کنید", "ابتدا یک ارائه‌دهنده انتخاب کنید", "پیش از ترجمه، یک ارائه‌دهنده هوش مصنوعی را در تنظیمات انتخاب کنید."],
  fi: ["Valitse palveluntarjoaja", "Valitse ensin palveluntarjoaja", "Valitse AI-palveluntarjoaja asetuksista ennen kääntämistä."],
  fr: ["Sélectionnez un fournisseur", "Sélectionnez d'abord un fournisseur", "Choisissez un fournisseur d'IA dans les paramètres avant de traduire."],
  he: ["בחרו ספק", "בחרו ספק תחילה", "בחרו ספק AI בהגדרות לפני התרגום."],
  hi: ["प्रदाता चुनें", "पहले प्रदाता चुनें", "अनुवाद करने से पहले सेटिंग में AI प्रदाता चुनें।"],
  hr: ["Odaberite pružatelja", "Najprije odaberite pružatelja", "Prije prevođenja odaberite pružatelja AI usluge u postavkama."],
  hu: ["Válasszon szolgáltatót", "Először válasszon szolgáltatót", "Fordítás előtt válasszon AI-szolgáltatót a beállításokban."],
  is: ["Veldu þjónustuaðila", "Veldu þjónustuaðila fyrst", "Veldu AI-þjónustuaðila í stillingum áður en þú þýðir."],
  it: ["Seleziona un provider", "Seleziona prima un provider", "Scegli un provider AI nelle impostazioni prima di tradurre."],
  ja: ["プロバイダーを選択", "先にプロバイダーを選択してください", "翻訳する前に、設定で AI プロバイダーを選択してください。"],
  ko: ["공급자 선택", "먼저 공급자를 선택하세요", "번역하기 전에 설정에서 AI 공급자를 선택하세요."],
  lv: ["Izvēlieties pakalpojumu sniedzēju", "Vispirms izvēlieties pakalpojumu sniedzēju", "Pirms tulkošanas iestatījumos izvēlieties AI pakalpojumu sniedzēju."],
  nl: ["Selecteer een provider", "Selecteer eerst een provider", "Kies een AI-provider in de instellingen voordat je vertaalt."],
  no: ["Velg en leverandør", "Velg en leverandør først", "Velg en AI-leverandør i innstillingene før du oversetter."],
  pl: ["Wybierz dostawcę", "Najpierw wybierz dostawcę", "Przed tłumaczeniem wybierz dostawcę AI w ustawieniach."],
  pt_BR: ["Selecione um provedor", "Selecione um provedor primeiro", "Escolha um provedor de IA nas configurações antes de traduzir."],
  pt_PT: ["Selecione um fornecedor", "Selecione primeiro um fornecedor", "Escolha um fornecedor de IA nas definições antes de traduzir."],
  ro: ["Selectați un furnizor", "Selectați mai întâi un furnizor", "Înainte de a traduce, alegeți un furnizor AI din setări."],
  ru: ["Выберите провайдера", "Сначала выберите провайдера", "Перед переводом выберите провайдера ИИ в настройках."],
  sl: ["Izberite ponudnika", "Najprej izberite ponudnika", "Pred prevajanjem v nastavitvah izberite ponudnika AI."],
  sr: ["Izaberite provajdera", "Prvo izaberite provajdera", "Pre prevođenja izaberite AI provajdera u podešavanjima."],
  sv: ["Välj en leverantör", "Välj en leverantör först", "Välj en AI-leverantör i inställningarna innan du översätter."],
  th: ["เลือกผู้ให้บริการ", "เลือกผู้ให้บริการก่อน", "เลือกผู้ให้บริการ AI ในการตั้งค่าก่อนแปล"],
  tr: ["Bir sağlayıcı seçin", "Önce bir sağlayıcı seçin", "Çeviri yapmadan önce ayarlardan bir yapay zekâ sağlayıcısı seçin."],
  ug: ["بىر تەمىنلىگۈچى تاللاڭ", "ئاۋۋال تەمىنلىگۈچى تاللاڭ", "تەرجىمە قىلىشتىن بۇرۇن تەڭشەكتىن بىر AI تەمىنلىگۈچىسىنى تاللاڭ."],
  uk: ["Виберіть постачальника", "Спочатку виберіть постачальника", "Перед перекладом виберіть постачальника ШІ в налаштуваннях."],
  vi: ["Chọn nhà cung cấp", "Chọn nhà cung cấp trước", "Chọn nhà cung cấp AI trong cài đặt trước khi dịch."],
};

const KEYS = Object.keys(EN);
const codes = fs
  .readdirSync(LOCALES)
  .filter((code) => fs.existsSync(path.join(LOCALES, code, "messages.json")))
  .sort();

const read = (code) => JSON.parse(fs.readFileSync(path.join(LOCALES, code, "messages.json"), "utf8"));

/** English text → translated text, harvested from the extension's catalogue. */
function translationMemory(code) {
  const extEnFile = path.join(EXT, "en", "messages.json");
  const extFile = path.join(EXT, code, "messages.json");
  if (!fs.existsSync(extEnFile) || !fs.existsSync(extFile)) return {};
  const extEn = JSON.parse(fs.readFileSync(extEnFile, "utf8"));
  const extLocale = JSON.parse(fs.readFileSync(extFile, "utf8"));
  const memory = {};
  for (const key of Object.keys(extEn)) {
    const source = extEn[key] && extEn[key].message;
    const target = extLocale[key] && extLocale[key].message;
    if (source && target) memory[source] = target;
  }
  return memory;
}

async function modelTranslate(languageName, source) {
  const res = await fetch(`${BASE}/chat/completions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: MODEL,
      temperature: 0.1,
      reasoning_effort: "none",
      messages: [
        { role: "system", content: "You are a careful software localizer. You answer with JSON only." },
        {
          role: "user",
          content: `Translate these UI strings into ${languageName}. Keep product names (Google, OpenAI, DualTran) and any placeholder unchanged. Output ONLY a JSON object with the same keys.\n${JSON.stringify(EN, null, 2)}`,
        },
      ],
    }),
    signal: AbortSignal.timeout(4000),
  });
  if (!res.ok) throw new Error(`model HTTP ${res.status}`);
  const data = await res.json();
  const text = ((data.choices || [])[0]?.message?.content || "").trim();
  return JSON.parse(text.slice(text.indexOf("{"), text.lastIndexOf("}") + 1));
}

async function main() {
  const report = [];
  for (const code of codes) {
    if (HANDWRITTEN.has(code)) continue;
    const catalogue = read(code);
    const missing = KEYS.filter((key) => !catalogue[key] || !String(catalogue[key].message || "").trim());
    if (!missing.length) continue;

    const memory = translationMemory(code);
    const resolved = {};
    for (const key of missing) if (memory[EN[key]]) resolved[key] = memory[EN[key]];

    const unresolved = missing.filter((key) => !resolved[key]);
    let model = {};
    if (unresolved.length) {
      try {
        model = await modelTranslate(new Intl.DisplayNames([code], { type: "language" }).of(code), Object.fromEntries(unresolved.map((key) => [key, EN[key]])));
      } catch (error) {
        model = {};
      }
    }
    for (const key of unresolved) {
      if (typeof model[key] === "string" && model[key].trim()) resolved[key] = model[key].trim();
    }

    const table = TABLE[code];
    const stillMissing = missing.filter((key) => !resolved[key]);
    if (stillMissing.length) {
      if (!table) {
        console.error(`fail ${code}: no translation source for ${stillMissing.join(", ")}`);
        process.exit(1);
      }
      // TABLE columns follow KEYS order.
      stillMissing.forEach((key) => { resolved[key] = table[KEYS.indexOf(key)]; });
    }
    for (const key of missing) catalogue[key] = { message: resolved[key] };
    fs.writeFileSync(path.join(LOCALES, code, "messages.json"), JSON.stringify(catalogue, null, 2) + "\n");
    report.push(`${code}: +${missing.length}`);
  }
  console.log(report.length ? report.join(", ") : "nothing to fill");
}

await main();

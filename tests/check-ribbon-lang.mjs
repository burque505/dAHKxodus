// Checks lang\*.ribbon.json against the ribbon dAHKxodus runs (web\node_modules\docxodus\dist\ribbon-chrome.js): every
// key must be a text the ribbon shows (else it is stale: the ribbon's English changed), and it lists the ribbon's texts
// still without a translation.
//   node tests\check-ribbon-lang.mjs
import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const { RIBBON_HTML: chrome } = await import(pathToFileURL(join(root, "web/node_modules/docxodus/dist/ribbon-chrome.js")).href);
// The ribbon's dialogs are part of it too (their preview's sample text is not translated).
const { PARAGRAPH_DIALOG_HTML } = await import(pathToFileURL(join(root, "web/node_modules/docxodus/dist/ribbon-paragraph-dialog.js")).href);
const { LIST_DIALOG_HTML, LIST_DIALOG_TEXTS } = await import(pathToFileURL(join(root, "web/node_modules/docxodus/dist/ribbon-list-dialog.js")).href);
// (the List Levels dialog writes its rows' choices from code: they are listed as texts of their own)
// (and the find bar's Special menu, made from code for each mode, and its "Invalid" for a pattern that isn't one)
const { specialFindCodes, specialReplaceCodes } = await import(pathToFileURL(join(root, "web/node_modules/docxodus/dist/find-query.js")).href);
const FIND_TEXTS = new Set(["Invalid"]);
for (const mode of [{}, { wildcards: true }, { regex: true }])
  for (const c of [...specialFindCodes(mode), ...specialReplaceCodes(mode)]) FIND_TEXTS.add(c.label);
const RIBBON_HTML = chrome + PARAGRAPH_DIALOG_HTML.replace(/<div class="dxr-dpreview"[\s\S]*?<\/div>/, "")
  + LIST_DIALOG_HTML + [...LIST_DIALOG_TEXTS, ...FIND_TEXTS].map((t) => `<span>${t}</span>`).join("");
const decode = (s) => s.replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&times;/g, "×").replace(/&amp;/g, "&")
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&nbsp;/g, " ");
const html = RIBBON_HTML.replace(/<!--[\s\S]*?-->/g, "").replace(/<style[\s\S]*?<\/style>/g, "").replace(/<svg[\s\S]*?<\/svg>/g, "");
const texts = new Set();
for (const m of html.matchAll(/>([^<]+)</g)) { const t = decode(m[1]).trim(); if (t) texts.add(t); }
for (const m of html.matchAll(/\s(?:title|aria-label|placeholder)="([^"]+)"/g)) texts.add(decode(m[1]).trim());
// (texts that read the same in any language: names, keys, number samples)
const words = [...texts].filter((t) => /[A-Za-z]{2}/.test(t) && !/^(Docxodus|Aa|pt|Ctrl\+\w|https:\/\/|[ivxIVX., ()]+|DOCX → WASM → DOCX)$/.test(t));

let failed = 0;
for (const file of readdirSync(join(root, "lang")).filter((f) => f.endsWith(".ribbon.json"))) {
  const labels = JSON.parse(readFileSync(join(root, "lang", file), "utf8"));
  const stale = Object.keys(labels).filter((k) => !texts.has(k));
  const missing = words.filter((t) => !(t in labels));
  for (const k of stale) { console.log(`FAILED ${file}: "${k}" is not a text the ribbon shows`); failed++; }
  console.log(`ok     ${file}: ${Object.keys(labels).length} texts, ${missing.length} of the ribbon's ${words.length} not translated`
    + (missing.length ? `: ${missing.map((t) => JSON.stringify(t)).join(", ")}` : ""));
}
console.log(failed ? `\n${failed} FAILED` : "\nPASSED");
process.exit(failed ? 1 : 0);

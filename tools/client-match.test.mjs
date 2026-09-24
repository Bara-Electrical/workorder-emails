#!/usr/bin/env node
// Regression tests for findClient() — run with: npm test
//
// index.js is a single module that starts an HTTP server and an email poll loop the
// moment it is imported, so importing it here would boot the app. Instead the three
// declarations under test are lifted out of the source text and evaluated in isolation.
// That keeps production code untouched while still testing the shipped source rather
// than a hand-copied paraphrase that could drift away from it.
//
// If index.js is refactored and a marker below no longer matches, this fails loudly with
// the marker it could not find — that is a prompt to update the marker, never to weaken
// the assertions.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const indexPath = path.join(here, "..", "index.js");
const src = fs.readFileSync(indexPath, "utf8").replace(/\r\n/g, "\n"); // a Windows checkout has CRLF; the markers below are LF

function grab(startMarker, endMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf(endMarker, i);
  if (j === -1) throw new Error(`no end marker "${endMarker}" after: ${startMarker}`);
  return src.slice(i, j);
}

const extracted = [
  grab("const CLIENT_NAME_MAP = {", "\n};\n") + "\n};\n",
  grab("const EMAIL_DOMAIN_MAP = {", "\n};\n") + "\n};\n",
  grab("function normaliseClientName(name)", "\n}\n") + "\n}\n",
  grab("function clientNameForms(realEstateName)", "\n}\n") + "\n}\n",
  grab("function clientNameCandidates(realEstateName)", "\n}\n") + "\n}\n",
  grab("async function findClient(realEstateName)", "\n// Full state names"),
].join("\n");

// The cache maps and the API fallback are module-level in index.js; stub them so the
// cache path under test runs exactly as it does in production. arofloGet returning no
// clients means a test that reaches the live-API fallback fails rather than silently
// passing through it.
const module_ = `
const clientCache = new Map();
const clientCacheNormalised = new Map();
const toArray = x => (Array.isArray(x) ? x : [x]);
const arofloGet = async () => ({ clients: [] });
${extracted}
export { findClient, clientCache, clientCacheNormalised, normaliseClientName, CLIENT_NAME_MAP, EMAIL_DOMAIN_MAP };
`;
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "client-match-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const m = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

// Real Aroflo client names, chosen to include the duplicate and junk records that make
// name normalisation risky — not just the happy path.
const CLIENTS = [
  "Rightly.Realestate",
  "L.J. Hooker Leederville/Kwinana",
  "LJ Hooker Victoria Park - Belmont",
  "A.D. Engineering International",
  "Michaelkeil.com",
  "R & P.M Buswell",
  "Rentwest Solutions",
  "Driven Property Group",
  "Professionals Armadale Real Estate WA",
  "Robin  Wright", "Robin Wright",              // duplicate owner records, double space
  "S Class Property Group", "S-Class Property Group",
  "cam .", "Cameron Burchell", "Cameron  Best", // junk single-name records
  "emma .", "Emma Smith",
  "Steven  .", "SDRE Steven Davis Real Estate", // work order 3378: junk card vs real agency
  "Scott .", "Regina .",
  "M Property", "M Property Management", "CC Property Advisory", // legal vs trading name
  "First National Swans Residential", "Oscar D'Souza Real Estate",
  "@realty (WA)", "Acton | Belle Property Rockingham",
  "Pro Property Group Real Estate", "ProProperty Group",
  // Client-not-found alerts of 21-23 Sep 2026, each with its closest Aroflo rival so the
  // aliases below are shown to be unambiguous rather than merely lucky.
  "Raine and Horne Landsdale", "Raine & Horne Midland",
  "Morgan & Hayes Real Estate", "Michael  Hayes",
  "Century 21 Grand Alliance", "Century 21 Coast Realty Mandurah",
  "Certainty Property WA", "Certainty Property Pty LTd",
  // Reached only by sending domain — the work order names the owner's company, never this.
  "Coronis Now WA", "Oscar D'Souza Real Estate",
];

for (const name of CLIENTS) m.clientCache.set(name.toLowerCase(), { clientname: name, clientid: name });
m.clientCacheNormalised.clear();
for (const c of m.clientCache.values()) {
  const key = m.normaliseClientName(c.clientname);
  if (!key) continue;
  const bucket = m.clientCacheNormalised.get(key);
  if (bucket) bucket.push(c); else m.clientCacheNormalised.set(key, [c]);
}

const CASES = [
  // The reported bug: Aroflo has "Rightly.Realestate", work orders say "Rightly Realestate".
  ["Rightly Realestate",            "Rightly.Realestate"],
  ["Rightly Real Estate",           "Rightly.Realestate"],
  ["Rightly.Realestate",            "Rightly.Realestate"],
  ["RIGHTLY REALESTATE",            "Rightly.Realestate"],
  ["Rightly Realestate Pty Ltd",    "Rightly.Realestate"],
  // Same shape, other live clients.
  ["AD Engineering International",  "A.D. Engineering International"],
  ["R & PM Buswell",                "R & P.M Buswell"],
  // Must not regress: exact and starts-with matching that already worked.
  ["Rentwest Solutions",            "Rentwest Solutions"],
  ["Driven Property Group Pty Ltd", "Driven Property Group"],
  ["LJ Hooker Victoria Park",       "LJ Hooker Victoria Park - Belmont"],
  // Must not regress: names colliding with duplicate/junk records under normalisation.
  ["Cameron Burchell",              "Cameron Burchell"],
  ["Emma Smith",                    "Emma Smith"],
  ["Robin Wright",                  "Robin Wright"],
  ["S Class Property Group",        "S Class Property Group"],
  // Genuinely ambiguous — must stay unmatched rather than pick one of the two.
  ["SClass Property Group",         null],
  // Work order 3378. No matching tier can reach "SDRE Steven Davis Real Estate" from the
  // name on the work order — the leading initials defeat all three — so findClient ALONE
  // must return nothing. Critically it must never fall back to the junk single-name card
  // "Steven  ." via the derived first-word candidate "Steven", which is the regression this
  // guards. In production CLIENT_NAME_MAP bridges the gap before findClient is called; that
  // path is covered by MAPPED_CASES below.
  ["Steven Davis Real Estate",      null],
  ["Scott Palmer Realty",           null],
  ["Regina Property Group",         null],
  // A genuinely single-word agency name is still a whole name, so it may normalise-match.
  ["Steven",                        "Steven  ."],
  // Client-not-found alert of 21 Sep 2026 that needed no alias: the card is an exact match,
  // and the exact tier has to beat the "Certainty Property Pty LTd" prefix rival for that to
  // hold. It failed live only because the card was created after the cache was loaded.
  ["Certainty Property WA",         "Certainty Property WA"],
];

// CLIENT_NAME_MAP is applied to the AI's extracted name BEFORE findClient sees it (see
// createArofloJob). Its keys must be lowercase to be found, so a mis-cased or padded key
// is silently dead — these cases exercise the map and the matcher together, as production
// does, rather than trusting the map by eye.
const MAPPED_CASES = [
  ["Steven Davis Real Estate", "SDRE Steven Davis Real Estate"],
  ["steven davis real estate", "SDRE Steven Davis Real Estate"],
  ["STEVEN DAVIS REAL ESTATE", "SDRE Steven Davis Real Estate"],
  // Client-not-found alerts of 8 Sep 2026.
  ["M Property Management Pty Ltd T/A CC Property Advisory Australia", "CC Property Advisory"],
  ["Drivengroup", "Driven Property Group"],
  ["Driven Group", "Driven Property Group"],
  ["First National Real Estate Swans Residential", "First National Swans Residential"],
  // Client-not-found alerts of 9-11 Sep 2026.
  ["At Realty (WA)", "@realty (WA)"],
  ["at realty wa", "@realty (WA)"],
  ["Acton Belle Property Rockingham & Baldivis", "Acton | Belle Property Rockingham"],
  ["Acton Belle Property Rockingham and Baldivis", "Acton | Belle Property Rockingham"],
  // Client-not-found alerts of 21-23 Sep 2026. "&" and "and" are not interchangeable once
  // normaliseClientName has stripped the ampersand, in either direction, and a legal name in
  // front of a trading name is not a prefix of it.
  ["Raine & Horne Landsdale", "Raine and Horne Landsdale"],
  ["raine & horne landsdale", "Raine and Horne Landsdale"],
  ["Morgan and Hayes Real Estate", "Morgan & Hayes Real Estate"],
  ["Grand Alliance Property Group Pty Ltd T/As Century 21 Grand Alliance", "Century 21 Grand Alliance"],
];

// Without its alias the legal name is genuinely ambiguous — it starts-with matches both
// "M Property" and "M Property Management" — so findClient alone must decline rather than
// pick one. This asserts the alias is doing the work, not a lucky fuzzy hit.
const UNMAPPED_MUST_DECLINE = [
  "M Property Management Pty Ltd T/A CC Property Advisory Australia",
  "Drivengroup",
  "First National Real Estate Swans Residential",
  "At Realty (WA)",
  "Acton Belle Property Rockingham & Baldivis",
  "Raine & Horne Landsdale",
  "Morgan and Hayes Real Estate",
  "Grand Alliance Property Group Pty Ltd T/As Century 21 Grand Alliance",
];

// EMAIL_DOMAIN_MAP is the last resort in createArofloJob: when the AI's extracted name finds
// no client, the SENDER's domain is looked up instead. It exists for agencies whose work
// orders put a third party where the agency name belongs — an owner's company, a conveyancer —
// which varies per property, so there is no name to alias. Its keys must be bare lowercase
// domains to be found, so a mis-cased or @-prefixed key is silently dead. These cases run the
// map and the matcher together, as production does.
const DOMAIN_CASES = [
  // 23 Sep 2026: named "Porcherealty Pty Ltd", the owner's company. 722 jobs on this card.
  ["Rica.Velez@coronis.com.au", "Coronis Now WA"],
  ["someone@CORONIS.COM.AU",    "Coronis Now WA"],
  ["pm@oscardsouza.com.au",     "Oscar D'Souza Real Estate"],
];

// The names these work orders actually carry must NOT match anything, or the domain fallback
// would never be reached and the job would land on whatever the owner's company matched.
const DOMAIN_ONLY_NAMES_MUST_DECLINE = [
  "Porcherealty Pty Ltd",
  "Bellerose Property Conveyancing",
];

let pass = 0;
const failures = [];

for (const [from, expected] of DOMAIN_CASES) {
  const domain = from.split("@")[1];
  const mapped = m.EMAIL_DOMAIN_MAP[domain?.toLowerCase()];
  const got = mapped ? (await m.findClient(mapped))?.clientname ?? null : null;
  if (got === expected) pass++;
  else failures.push(`  ${JSON.stringify(from)} (via EMAIL_DOMAIN_MAP)\n    expected: ${expected}\n    got:      ${got}`);
}

for (const name of DOMAIN_ONLY_NAMES_MUST_DECLINE) {
  const mapped = m.CLIENT_NAME_MAP[name.toLowerCase()] || name;
  const got = (await m.findClient(mapped))?.clientname ?? null;
  if (got === null) pass++;
  else failures.push(`  ${JSON.stringify(name)} (should reach the domain fallback)\n    expected: null\n    got:      ${got}`);
}

// A key that is not a bare lowercase domain can never be found by the lookup above.
for (const key of Object.keys(m.EMAIL_DOMAIN_MAP)) {
  if (key === key.toLowerCase().trim() && !key.includes("@") && !key.includes("/") && key.includes(".")) pass++;
  else failures.push(`  EMAIL_DOMAIN_MAP key ${JSON.stringify(key)} is not a bare lowercase domain — it can never match`);
}
for (const [input, expected] of MAPPED_CASES) {
  const mapped = m.CLIENT_NAME_MAP[input.toLowerCase()] || input;
  const got = (await m.findClient(mapped))?.clientname ?? null;
  if (got === expected) pass++;
  else failures.push(`  ${JSON.stringify(input)} (via CLIENT_NAME_MAP)\n    expected: ${expected}\n    got:      ${got}`);
}

for (const [input, expected] of CASES) {
  const got = (await m.findClient(input))?.clientname ?? null;
  if (got === expected) pass++;
  else failures.push(`  ${JSON.stringify(input)}\n    expected: ${expected}\n    got:      ${got}`);
}

for (const input of UNMAPPED_MUST_DECLINE) {
  const got = (await m.findClient(input))?.clientname ?? null;
  if (got === null) pass++;
  else failures.push(`  ${JSON.stringify(input)} (unmapped)\n    expected: null\n    got:      ${got}`);
}

const total = CASES.length + MAPPED_CASES.length + UNMAPPED_MUST_DECLINE.length
  + DOMAIN_CASES.length + DOMAIN_ONLY_NAMES_MUST_DECLINE.length + Object.keys(m.EMAIL_DOMAIN_MAP).length;
console.log(`findClient: ${pass}/${total} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n${failures.join("\n")}`);
  process.exit(1);
}

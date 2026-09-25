#!/usr/bin/env node
// Regression tests for the site-equipment line in buildDescription() — run with: npm test
//
// index.js starts an HTTP server and the poll loop the moment it is imported, so the
// functions under test are lifted out of the source text and evaluated in isolation, the same
// way as the other tools/*.test.mjs files. templates.js is a plain module, so the real package
// templates are imported rather than stubbed.
//
// The dashboard's property check supplies an aircon tally for the site ("Site: 1× Evaporative").
// It only belongs on an aircon job: job 108163 was a bedroom light fitting that carried the
// property's evap unit in its description.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, "..");
const src  = fs.readFileSync(path.join(root, "index.js"), "utf8").replace(/\r\n/g, "\n");

function grabFn(startMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf("\n}\n", i);
  if (j === -1) throw new Error(`no end of function after: ${startMarker}`);
  return src.slice(i, j + 3);
}

const module_ = [
  `import { PACKAGE_TEMPLATES } from ${JSON.stringify(pathToFileURL(path.join(root, "templates.js")).href)};`,
  `import { siteLine } from ${JSON.stringify(pathToFileURL(path.join(root, "gate.js")).href)};`,
  grabFn("function escapeHtml(value)"),
  grabFn("function extractLockboxDetails(accessDetails)"),
  grabFn("function isAirconJob(result)"),
  grabFn("function tallyFromUnitTags(airconUnitType)"),
  grabFn("function reconcileAircon(airconUnitType, siteAircon)"),
  grabFn("function buildDescription(result, airconUnitType = null, site = \"\")"),
  "export { buildDescription, reconcileAircon, tallyFromUnitTags };",
].join("\n");

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "description-site-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const { buildDescription, reconcileAircon, tallyFromUnitTags } = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const SITE = "Site: 1× Evaporative";
const results = [];
const check = (name, cond) => results.push({ name, ok: !!cond });
const hasSite = html => html.includes(SITE);

// The live case: 108163, a light-fitting repair at a property with an evap unit on record.
{
  const html = buildDescription(
    { "task-type": "Real Estate General Maintenance", "package": null,
      "task-description": "Bedroom light socket/fitting has come away; inspect and repair." },
    null, SITE);
  check("108163: a non-aircon job does not carry the site's aircon tally", !hasSite(html));
  check("108163: the job's own description is still there", html.includes("Bedroom light socket"));
  // The spacer before the highlights must not be left dangling with nothing after it.
  check("108163: no orphan spacer when the site line is the only highlight", !html.trimEnd().endsWith("<p>&nbsp;</p>"));
}

// An electrical compliance job is not an aircon job either.
{
  const html = buildDescription({ "task-type": "EC1", "package": "EC1", "task-description": "Annual compliance" }, null, SITE);
  check("EC1 does not carry the aircon tally", !hasSite(html));
}

// Every aircon job still gets it.
for (const pkg of ["AC1", "AC2", "ACEC1"]) {
  const html = buildDescription({ "task-type": "Real Estate General Maintenance", "package": pkg, "task-description": "Service" }, null, SITE);
  check(`${pkg} carries the aircon tally`, hasSite(html));
}
{
  const html = buildDescription({ "task-type": "Real Estate Aircon Maintenance", "package": null, "task-description": "Not cooling" }, null, SITE);
  check("an aircon repair (no package) carries the aircon tally", hasSite(html));
}

// Nothing on record means nothing to show, on any job.
{
  const html = buildDescription({ "task-type": "Real Estate Aircon Maintenance", "task-description": "Not cooling" }, null, "");
  check("no tally on record, nothing shown", !html.includes("Site:"));
}

// The other highlights are untouched by the gate.
{
  const html = buildDescription(
    { "task-type": "Real Estate General Maintenance", "task-description": "Fix light",
      "expenditure-limit": "$330", "access-details": "Lockbox code: 214" },
    null, SITE);
  check("expenditure limit still shown on a non-aircon job", html.includes("Expenditure Limit: $330"));
  check("lockbox still shown on a non-aircon job", html.includes("Access Details: Lockbox code: 214"));
  check("…and the aircon tally still is not", !hasSite(html));
}

// ---- One aircon line, and whether its two sources agree ----
const eq = (name, got, expected) => check(`${name} (got ${JSON.stringify(got)})`, JSON.stringify(got) === JSON.stringify(expected));

eq("tag: single split",            tallyFromUnitTags("Split System"),          { Split: 1, Ducted: 0, Evaporative: 0 });
eq("tag: split x3",                tallyFromUnitTags("Split System x3"),       { Split: 3, Ducted: 0, Evaporative: 0 });
eq("tag: evap",                    tallyFromUnitTags("Evap"),                  { Split: 0, Ducted: 0, Evaporative: 1 });
eq("tag: two tags together",       tallyFromUnitTags("Split System x2, Ducted"), { Split: 2, Ducted: 1, Evaporative: 0 });
eq("tag: none",                    tallyFromUnitTags(null),                    { Split: 0, Ducted: 0, Evaporative: 0 });

const cases = [
  // [name, tag, site record, expected unitType, expected site line, expected mismatch]
  ["match: evap tagged, evap on record",            "Evap",                    { Evaporative: 1 },         "Evap",  "", false],
  ["match: split x2 tagged, 2 splits on record",    "Split System x2",         { Split: 2 },               "Split System x2", "", false],
  ["match: two tags, both on record",               "Split System x2, Ducted", { Split: 2, Ducted: 1 },    "Split System x2, Ducted", "", false],
  ["match: dashboard's { units } wrapper",          "Evap",                    { units: { Evaporative: 1 } }, "Evap", "", false],
  ["mismatch: evap tagged, splits on record",       "Evap",                    { Split: 2 },               "Evap",  "", true],
  ["mismatch: right type, wrong count",             "Split System x2",         { Split: 3 },               "Split System x2", "", true],
  ["mismatch: tag misses a unit on record",         "Split System",            { Split: 1, Evaporative: 1 }, "Split System", "", true],
  ["tag only, nothing on record — no comparison",   "Evap",                    {},                         "Evap",  "", false],
  ["record only, no tag — record is the line",      null,                      { Evaporative: 1 },         null,    "Site: 1× Evaporative", false],
  ["neither — nothing at all",                      null,                      {},                         null,    "", false],
];
for (const [name, tag, site, unitType, siteText, mismatch] of cases) {
  const r = reconcileAircon(tag, site);
  check(`${name}: line`,     r.unitType === unitType && r.site === siteText);
  check(`${name}: mismatch`, r.mismatch === mismatch);
  check(`${name}: never both lines`, !(r.unitType && r.site));
}
check("mismatch names what the record says", reconcileAircon("Evap", { Split: 2 }).siteSummary === "2× Split");

// End to end: an aircon job with both a tag and a site record carries exactly one aircon line.
{
  const job = { "task-type": "Real Estate Aircon Maintenance", "package": "AC1", "task-description": "Service" };
  const r = reconcileAircon("Evap", { Evaporative: 1 });
  const html = buildDescription(job, r.unitType, r.site);
  check("aircon job with both sources shows the tag", html.includes("Unit Type: Evap"));
  check("…and not the site record as well", !html.includes("Site:"));
}
{
  const job = { "task-type": "Real Estate Aircon Maintenance", "package": "AC1", "task-description": "Service" };
  const r = reconcileAircon(null, { Evaporative: 1 });
  const html = buildDescription(job, r.unitType, r.site);
  check("aircon job with no tag falls back to the site record", html.includes("Site: 1× Evaporative") && !html.includes("Unit Type"));
}

const failures = results.filter(r => !r.ok);
console.log(`description site line: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f => `  ${f.name}`).join("\n"));
  process.exit(1);
}

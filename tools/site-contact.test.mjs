#!/usr/bin/env node
// Regression tests for the lockbox-as-site-contact rule — run with: npm test
//
// index.js starts an HTTP server and the poll loop the moment it is imported, so the
// functions under test are lifted out of the source text and evaluated in isolation, the same
// way as the other tools/*.test.mjs files.
//
// The office's rule: a work order with no tenant details means the tenant no longer lives
// there, so the previous tenant's name, phone and email are cleared — and when there is a
// lockbox code it takes the contact slot (job 108212 was filled in by hand). A named tenant,
// "Vacant" included, is written as before.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const src  = fs.readFileSync(path.join(here, "..", "index.js"), "utf8").replace(/\r\n/g, "\n");

function grabFn(startMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  const j = src.indexOf("\n}\n", i);
  if (j === -1) throw new Error(`no end of function after: ${startMarker}`);
  return src.slice(i, j + 3);
}
function grabLine(startMarker) {
  const i = src.indexOf(startMarker);
  if (i === -1) throw new Error(`index.js no longer contains: ${startMarker}`);
  return src.slice(i, src.indexOf("\n", i) + 1);
}

const module_ = [
  grabLine("const SITE_FIELD_LIMIT"),
  grabFn("function extractLockboxDetails(accessDetails)"),
  grabFn("function lockboxSiteContact(tenantName, accessDetails)"),
  grabFn("function locationContactUpdate(tenantName, tenantContact, tenantEmail, lockboxContact)"),
  "export { lockboxSiteContact, locationContactUpdate, SITE_FIELD_LIMIT };",
].join("\n");

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "site-contact-"));
const tmp = path.join(tmpDir, "extracted.mjs");
fs.writeFileSync(tmp, module_);
const m = await import(pathToFileURL(tmp).href);
process.on("exit", () => fs.rmSync(tmpDir, { recursive: true, force: true }));

const results = [];
const eq = (name, got, expected) =>
  results.push({ name, ok: JSON.stringify(got) === JSON.stringify(expected), got, expected });

// ---- lockboxSiteContact ----
const LB = "Lockbox code: 214";
eq("108212: no tenant, lockbox",               m.lockboxSiteContact(null, LB),                        LB);
eq("a named tenant takes the slot",            m.lockboxSiteContact("Emma", LB),                      null);
eq("'Vacant' takes the slot",                  m.lockboxSiteContact("Vacant", LB),                    null);
eq("whitespace-only tenant counts as none",    m.lockboxSiteContact("  ", LB),                        LB);
eq("no lockbox, nothing to add",               m.lockboxSiteContact(null, "Key: 12, Gate code: 9"),  null);
eq("no access details at all",                 m.lockboxSiteContact(null, null),                      null);
eq("lockbox picked out of mixed details",      m.lockboxSiteContact(null, "Key: 12, Lockbox code: 7"), "Lockbox code: 7");
{
  const long = "Lockbox code: " + "9".repeat(80);
  eq("capped at AroFlo's site-contact limit",  m.lockboxSiteContact(null, long).length,               m.SITE_FIELD_LIMIT);
}

// ---- locationContactUpdate: null = leave alone, "" = clear ----
// No tenant details at all: the previous tenant has moved out, so everything of theirs goes.
eq("108212: no tenant details, lockbox — contact is the lockbox, old details cleared",
   m.locationContactUpdate(null, undefined, undefined, LB),
   { sitecontact: LB, sitephone: "", siteemail: "" });
eq("no tenant details, no lockbox — old tenant cleared entirely",
   m.locationContactUpdate(null, undefined, undefined, null),
   { sitecontact: "", sitephone: "", siteemail: "" });
eq("empty strings count as no details",
   m.locationContactUpdate("", "", "", null),
   { sitecontact: "", sitephone: "", siteemail: "" });

// A named tenant is the current state and clears whatever of the old tenant it doesn't replace.
eq("named tenant with their own details",
   m.locationContactUpdate("Emma", "0400 000 000", "e@x.com", null),
   { sitecontact: "Emma", sitephone: "0400 000 000", siteemail: "e@x.com" });
eq("named tenant clears a stale phone and email",
   m.locationContactUpdate("Emma", undefined, undefined, null),
   { sitecontact: "Emma", sitephone: "", siteemail: "" });
eq("'Vacant' clears the old tenant's details",
   m.locationContactUpdate("Vacant", undefined, undefined, null),
   { sitecontact: "Vacant", sitephone: "", siteemail: "" });
eq("a named tenant wins over a lockbox",
   m.locationContactUpdate("Emma", undefined, undefined, LB),
   { sitecontact: "Emma", sitephone: "", siteemail: "" });

// A phone or email without a name is partial tenant info: write it, leave the rest.
eq("phone but no name — phone written, name left",
   m.locationContactUpdate(null, "0412 345 678", undefined, LB),
   { sitecontact: null, sitephone: "0412 345 678", siteemail: null });
eq("email but no name — email written, rest left",
   m.locationContactUpdate(null, undefined, "t@x.com", null),
   { sitecontact: null, sitephone: null, siteemail: "t@x.com" });

const failures = results.filter(r => !r.ok);
console.log(`site contact: ${results.length - failures.length}/${results.length} passed`);
if (failures.length) {
  console.error(`\n${failures.length} failure(s):\n` + failures.map(f =>
    `  ${f.name}\n    expected: ${JSON.stringify(f.expected)}\n    got:      ${JSON.stringify(f.got)}`).join("\n"));
  process.exit(1);
}

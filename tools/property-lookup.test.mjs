#!/usr/bin/env node
// Tests for the evap / two-storey property check (propertyLookup.js) — run with: npm test

import assert from "node:assert/strict";
import { parseLookup, lookupFindings, fetchRoofPhoto, landgateUrl, hereSatelliteUrl, lookupProperty } from "../propertyLookup.js";

let failed = 0;
async function test(name, fn) {
  try { await fn(); console.log(`  ok  ${name}`); }
  catch (err) { failed++; console.log(`  FAIL ${name}\n       ${err.message}`); }
}

const json = (o) => JSON.stringify(o);
const listing = (o = {}) => ({ found: true, url: "https://www.realestate.com.au/property-house-wa-leeming-1", evap: "unknown", storeys: "unknown", evidence: "", ...o });
const roof = (o = {}) => ({ usable: true, evap: "unknown", confidence: 0, evidence: "", ...o });
const lookup = (l, r) => parseLookup(json({ listing: listing(l), roof: roof(r) }));

await test("parse: junk and missing fields read as unknown", () => {
  assert.equal(parseLookup("not json"), null);
  const p = parseLookup(`Here you go: ${json({ listing: { found: true, evap: "maybe", storeys: 3, url: "javascript:x" } })}`);
  assert.deepEqual(p.listing, { found: true, url: null, evap: "unknown", storeys: "unknown", evidence: "" });
  assert.equal(p.roof.evap, "unknown");
});

await test("parse: no matching listing means nothing from the listing counts", () => {
  const p = parseLookup(json({ listing: { found: false, evap: "yes", storeys: 2, url: "https://x.au" }, roof: roof() }));
  assert.equal(p.listing.evap, "unknown");
  assert.equal(p.listing.storeys, "unknown");
  assert.equal(p.listing.url, null);
});

await test("parse: an unusable photo never says evap", () => {
  const p = parseLookup(json({ listing: listing(), roof: { usable: false, evap: "yes", confidence: 0.9 } }));
  assert.equal(p.roof.evap, "unknown");
});

await test("AC1 with evap in the listing: Evap found, no task type change", () => {
  const f = lookupFindings(lookup({ evap: "yes", evidence: "ducted evaporative cooling" }), "AC1");
  assert.deepEqual(f.warnings.map(w => w.tag), ["Evap found"]);
  assert.match(f.warnings[0].detail, /ducted evaporative cooling/);
  assert.match(f.line.text, /evap/);
});

await test("AC1 with a sure roof photo: Evap found; an unsure one: nothing", () => {
  assert.deepEqual(lookupFindings(lookup({}, { evap: "yes", confidence: 0.8 }), "ACEC1").warnings.map(w => w.tag), ["Evap found"]);
  const weak = lookupFindings(lookup({}, { evap: "yes", confidence: 0.5 }), "ACEC1");
  assert.deepEqual(weak, { warnings: [], line: null });
});

await test("Two storey only flags evap jobs", () => {
  assert.deepEqual(lookupFindings(lookup({ storeys: 2 }), "AC1").warnings, []);
  assert.deepEqual(lookupFindings(lookup({ storeys: 2 }), "EV1").warnings.map(w => w.tag), ["Two storey"]);
  assert.deepEqual(lookupFindings(lookup({ storeys: 2, evap: "yes" }), "AC2").warnings.map(w => w.tag), ["Evap found", "Two storey"]);
});

await test("EV1 already: evap found is no warning, just the line", () => {
  const f = lookupFindings(lookup({ evap: "yes", storeys: 1 }), "EVEC1");
  assert.deepEqual(f.warnings, []);
  assert.equal(f.line.text, "Property check (listing + roof photo): evap, single storey");
  assert.equal(f.line.url, "https://www.realestate.com.au/property-house-wa-leeming-1");
});

await test("all unknown: no warnings, no line", () => {
  assert.deepEqual(lookupFindings(lookup({}, {}), "AC1"), { warnings: [], line: null });
  assert.deepEqual(lookupFindings(null, "AC1"), { warnings: [], line: null });
});

const img = (type = "image/jpeg", bytes = 20000) => new Response(Buffer.alloc(bytes, 1), { status: 200, headers: { "content-type": type } });

await test("roof photo: Landgate first, HERE when Landgate fails, none when both fail", async () => {
  const seen = [];
  let p = await fetchRoofPhoto({ lat: -32.07, lon: 115.87, hereApiKey: "k", fetchImpl: async (u) => { seen.push(u); return img(); } });
  assert.equal(p.source, "Landgate");
  assert.equal(seen.length, 1);
  p = await fetchRoofPhoto({ lat: -32.07, lon: 115.87, hereApiKey: "k", fetchImpl: async (u) => u.includes("slip.wa.gov.au") ? new Response("{\"error\":1}", { headers: { "content-type": "application/json" } }) : img() });
  assert.equal(p.source, "HERE");
  p = await fetchRoofPhoto({ lat: -32.07, lon: 115.87, hereApiKey: null, fetchImpl: async () => img("image/jpeg", 100) });
  assert.equal(p, null);
});

await test("image URLs", () => {
  assert.match(landgateUrl(-32.07, 115.87), /^https:\/\/services\.slip\.wa\.gov\.au\/.*export\?bbox=115\.86\d+%2C-32\.07\d+/);
  assert.equal(hereSatelliteUrl(-32.07, 115.87, "a b"), "https://image.maps.hereapi.com/mia/v3/base/mc/center:-32.07,115.87;zoom=20/1024x1024/jpeg?style=satellite.day&apiKey=a%20b");
});

await test("lookupProperty sends the photo and web search, and never throws", async () => {
  let req = null;
  const openai = { responses: { create: async (r) => { req = r; return { output_text: json({ listing: listing({ evap: "yes" }), roof: roof() }) }; } } };
  const out = await lookupProperty({ openai, address: "13 Devenish Way, Leeming WA 6149", coords: { lat: -32.07, lon: 115.87 }, fetchImpl: async () => img() });
  assert.equal(out.listing.evap, "yes");
  assert.equal(out.photoSource, "Landgate");
  assert.equal(req.tools[0].type, "web_search");
  assert.equal(req.input[0].content[1].type, "input_image");
  const broken = { responses: { create: async () => { throw new Error("429"); } } };
  assert.equal(await lookupProperty({ openai: broken, address: "x", coords: null }), null);
});

if (failed) { console.log(`\n${failed} failed`); process.exit(1); }
console.log("\nproperty lookup: all passed");

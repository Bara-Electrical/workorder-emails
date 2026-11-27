#!/usr/bin/env node
// Tests for the work order photo check (photoCheck.js) — run with: npm test

import assert from "node:assert/strict";
import { parseFindings, photoActions, photoCheckMode, photoEmail, checkPhotos, WRITE_CONFIDENCE } from "../photoCheck.js";

let failed = 0;
async function test(name, fn) {
  try { await fn(); console.log(`  ok  ${name}`); }
  catch (err) { failed++; console.log(`  FAIL ${name}\n       ${err.message}`); }
}

const f = (o = {}) => ({ kind: "appliance_90cm", detail: "90cm Westinghouse cooktop", photos: [1], confidence: 0.95, evidence: "rating plate reads WHC943SC", ...o });
const json = (findings) => JSON.stringify({ findings });

await test("mode: email unless set to off or aroflo", () => {
  assert.equal(photoCheckMode({}), "email");
  assert.equal(photoCheckMode({ PHOTO_CHECK: "AroFlo " }), "aroflo");
  assert.equal(photoCheckMode({ PHOTO_CHECK: "off" }), "off");
  assert.equal(photoCheckMode({ PHOTO_CHECK: "on" }), "email");
});

await test("parse: junk is null, bad findings are dropped", () => {
  assert.equal(parseFindings("no json here"), null);
  const p = parseFindings(json([
    f(),
    f({ kind: "pool_pump" }),
    f({ evidence: "" }),
    f({ confidence: 1.4 }),
    f({ kind: "aircon", aircon_type: "fridge", detail: "1x fridge" }),
    f({ kind: "aircon", aircon_type: "Ducted", aircon_count: 1, detail: "1x ducted" }),
    f({ photos: [0, 2, 9] }),
  ]), 3);
  assert.equal(p.findings.length, 3);
  assert.equal(p.findings[1].airconType, "ducted");
  assert.deepEqual(p.findings[2].photos, [2]);
});

await test("actions: a clear 90cm finding goes to the scheduling note", () => {
  const a = photoActions(parseFindings(json([f()])).findings, { text: "Cooktop not working", taskType: "Real Estate General Maintenance" });
  assert.deepEqual(a.schedulingLines, ["Photos: 90cm Westinghouse cooktop - 90cm, 2 people"]);
  assert.equal(a.descriptionLine, null);
});

await test("actions: below the bar, nothing is written", () => {
  const a = photoActions([f({ confidence: WRITE_CONFIDENCE - 0.01 })], { text: "" });
  assert.deepEqual(a.schedulingLines, []);
  assert.equal(a.skipped[0].reason, "not sure enough");
});

await test("actions: nothing the work order already says", () => {
  const a = photoActions([f(), f({ kind: "double_oven", detail: "stacked double oven" })], { text: "Replace the 900mm cooktop. Double oven also faulty." });
  assert.deepEqual(a.schedulingLines, []);
  assert.equal(a.skipped.length, 2);
});

await test("actions: two storey only matters on roof work", () => {
  const ts = f({ kind: "two_storey", detail: "two storey house" });
  assert.deepEqual(photoActions([ts], { text: "Replace kitchen light", taskType: "Real Estate General Maintenance" }).schedulingLines, []);
  assert.deepEqual(photoActions([ts], { text: "Service the evap", pkg: "EV1" }).schedulingLines, ["Photos: two storey - 2 people for the roof"]);
});

await test("actions: aircon line only on aircon jobs with no unit type known, one line", () => {
  const ac = f({ kind: "aircon", detail: "2x split", airconType: "split", airconCount: 2 });
  const ac2 = f({ kind: "aircon", detail: "1x ducted", airconType: "ducted", airconCount: 1 });
  assert.equal(photoActions([ac, ac2], { pkg: "AC1" }).descriptionLine, "Photos: 2× Split");
  assert.equal(photoActions([ac], { pkg: "AC1", airconKnown: true }).descriptionLine, null);
  assert.equal(photoActions([ac], { taskType: "Real Estate General Maintenance" }).descriptionLine, null);
  assert.equal(photoActions([ac], { taskType: "Real Estate Aircon Maintenance" }).descriptionLine, "Photos: 2× Split");
});

await test("actions: high ceiling asks for a tall ladder", () => {
  const a = photoActions([f({ kind: "high_ceiling", detail: "raked ceiling about 4.5m at the light" })], { text: "Light not working" });
  assert.deepEqual(a.schedulingLines, ["Photos: raked ceiling about 4.5m at the light - tall ladder"]);
});

await test("email: only when something was found, says whether it was written", () => {
  const check = { findings: [f()], photoCount: 2 };
  const actions = photoActions(check.findings, { text: "" });
  const off = photoEmail({ jobNumber: "111500", address: "1 Test St", check, actions, written: false });
  assert.ok(off.worthSending);
  assert.match(off.html, /Would add to the job/);
  const on = photoEmail({ jobNumber: "111500", address: "1 Test St", check, actions, written: true });
  assert.match(on.html, /Added to the job/);
  assert.equal(photoEmail({ jobNumber: "1", check: { findings: [], photoCount: 1 }, actions: photoActions([], {}) }).worthSending, false);
});

await test("check: no photos or no client means no call", async () => {
  assert.equal(await checkPhotos({ openai: null, images: [{ data: new Uint8Array(1) }] }), null);
  let called = false;
  const openai = { responses: { create: async () => { called = true; return { output_text: "{}" }; } } };
  assert.equal(await checkPhotos({ openai, images: [] }), null);
  assert.equal(called, false);
});

await test("check: sends at most 8 photos and reads the answer; a failure is null", async () => {
  let sent = 0;
  const openai = { responses: { create: async (req) => {
    sent = req.input[0].content.filter(c => c.type === "input_image").length;
    return { output_text: json([f({ photos: [9] })]) };
  } } };
  const images = Array.from({ length: 12 }, (_, i) => ({ name: `p${i}.jpg`, data: new Uint8Array([1, 2, 3]), contentType: "image/jpeg" }));
  const r = await checkPhotos({ openai, images });
  assert.equal(sent, 8);
  assert.equal(r.photoCount, 8);
  assert.deepEqual(r.findings[0].photos, []);
  const broken = { responses: { create: async () => { throw new Error("timeout"); } } };
  assert.equal(await checkPhotos({ openai: broken, images }), null);
});

if (failed) { console.log(`\n${failed} failed`); process.exit(1); }
console.log("\nphoto check: all passed");

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  gateMode, gateApplies, siteLine, propertyCheck, upsertGateRecord, pendingDecisions, GateHold,
  GATE_TEST_CATEGORY, GATE_CONTINUE_CATEGORY,
} from "./gate.js";

const env = { DASHBOARD_URL: "https://dash.test/", DASHBOARD_API_SECRET: "s3cret" };

// A scripted fetch: each call pops the next reply; the log records what was asked.
function scripted(replies) {
  const log = [];
  const f = async (url, init = {}) => {
    log.push({ url: String(url), method: init.method ?? "GET", headers: init.headers ?? {}, body: init.body });
    const r = replies.shift();
    if (!r) throw new Error(`unexpected fetch ${url}`);
    return { ok: r.status < 300, status: r.status, text: async () => r.body ?? "", json: async () => JSON.parse(r.body) };
  };
  return { f, log };
}

test("gateMode defaults to off and falls back to off on an unknown value", () => {
  assert.equal(gateMode({}), "off");
  assert.equal(gateMode({ WORKORDER_GATE: "" }), "off");
  assert.equal(gateMode({ WORKORDER_GATE: "tagged" }), "tagged");
  assert.equal(gateMode({ WORKORDER_GATE: " ON " }), "on");
  assert.equal(gateMode({ WORKORDER_GATE: "maybe" }), "off");
});

test("gateApplies: off never, tagged only with Gate Test, on always — Gate: Continue disables all", () => {
  const plain = ["Bara AI"], tagged = ["Bara AI", GATE_TEST_CATEGORY], cont = [GATE_CONTINUE_CATEGORY];
  const matrix = [
    ["off",    plain,  false], ["off",    tagged, false], ["off",    [...tagged, ...cont], false],
    ["tagged", plain,  false], ["tagged", tagged, true],  ["tagged", [...tagged, ...cont], false],
    ["on",     plain,  true],  ["on",     tagged, true],  ["on",     [...plain, ...cont],  false],
  ];
  for (const [mode, cats, want] of matrix) assert.equal(gateApplies(mode, cats), want, `${mode} × ${cats.join("+")}`);
  assert.equal(gateApplies("on", undefined), true);
});

test("siteLine orders Split, Ducted, Evaporative, omits zeros, and is empty with nothing", () => {
  assert.equal(siteLine({ Evaporative: 1, Split: 2, Ducted: 1 }), "Site: 2× Split, 1× Ducted, 1× Evaporative");
  assert.equal(siteLine({ Split: 2, Ducted: 0 }), "Site: 2× Split");
  assert.equal(siteLine({ units: { Ducted: 1 } }), "Site: 1× Ducted");
  assert.equal(siteLine({}), "");
  assert.equal(siteLine(null), "");
});

test("propertyCheck GETs with the bearer and returns the body", async () => {
  const body = { locationId: "L1", recentJobs: [{ jobNumber: "107620" }], aircon: { units: { Split: 1 } } };
  const { f, log } = scripted([{ status: 200, body: JSON.stringify(body) }]);
  assert.deepEqual(await propertyCheck({ aroFloLocationId: "JiQq==", clientAroFloId: "C1", street: "2 Nardoo Way", suburb: "Maddington" }, { fetchImpl: f, env }), body);
  assert.equal(log[0].url, "https://dash.test/api/property-check?aroFloLocationId=JiQq%3D%3D&clientAroFloId=C1&street=2+Nardoo+Way&suburb=Maddington");
  assert.equal(log[0].headers.Authorization, "Bearer s3cret");
});

test("propertyCheck throws on a non-200 and when the dashboard is not configured", async () => {
  const { f } = scripted([{ status: 503, body: "down" }]);
  await assert.rejects(propertyCheck({ aroFloLocationId: "x" }, { fetchImpl: f, env }), /Property check failed: 503 down/);
  await assert.rejects(propertyCheck({ aroFloLocationId: "x" }, { fetchImpl: f, env: {} }), /DASHBOARD_URL/);
});

test("upsertGateRecord PUTs the fields as JSON with the bearer", async () => {
  const { f, log } = scripted([{ status: 200, body: "{}" }]);
  await upsertGateRecord({ messageId: "m1", status: "CHECKED" }, { fetchImpl: f, env });
  assert.equal(log[0].method, "PUT");
  assert.equal(log[0].url, "https://dash.test/api/workorder-gate");
  assert.equal(log[0].headers.Authorization, "Bearer s3cret");
  assert.deepEqual(JSON.parse(log[0].body), { messageId: "m1", status: "CHECKED" });
});

test("pendingDecisions reads the dashboard's { items } and tolerates an empty body", async () => {
  const { f, log } = scripted([
    { status: 200, body: JSON.stringify({ items: [{ messageId: "a" }] }) },
    { status: 200, body: JSON.stringify({}) },
  ]);
  assert.deepEqual(await pendingDecisions("CONTINUED", { fetchImpl: f, env }), [{ messageId: "a" }]);
  assert.deepEqual(await pendingDecisions("STOPPED", { fetchImpl: f, env }), []);
  assert.equal(log[0].url, "https://dash.test/api/workorder-gate/pending?status=CONTINUED");
});

test("GateHold is an Error carrying the checks", () => {
  const hold = new GateHold({ recentJobs: [{ jobNumber: "1" }], aircon: { Split: 1 } });
  assert.ok(hold instanceof Error);
  assert.equal(hold.recentJobs.length, 1);
  assert.deepEqual(hold.aircon, { Split: 1 });
  assert.match(hold.message, /1 recent job/);
});

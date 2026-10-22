import { test } from "node:test";
import assert from "node:assert/strict";
import { buildPickInput, parsePick, aiPickClient, closestClients } from "./clientPick.js";

const CARDS = [
  { clientid: "a", clientname: "Ray White South Perth" },
  { clientid: "b", clientname: "Ray White ELG" },
  { clientid: "c", clientname: "Cooling Real Estate" },
  { clientid: "d", clientname: "Morgan & Hayes Real Estate" },
  { clientid: "e", clientname: "Smith Real Estate" },
];

test("parsePick takes only a card name from the list, case and spacing forgiven", () => {
  assert.equal(parsePick('{"card":"cooling  real estate","reason":"letterhead"}', CARDS).client.clientid, "c");
  assert.equal(parsePick('{"card":"Cooling Real Estate","reason":"x"}', CARDS).reason, "x");
});

test("parsePick: null, a near miss, or bad JSON is no answer", () => {
  assert.equal(parsePick('{"card":null,"reason":"new agency"}', CARDS), null);
  assert.equal(parsePick('{"card":"Cooling Realty","reason":""}', CARDS), null);
  assert.equal(parsePick("not json", CARDS), null);
});

test("parsePick: two cards with the same name is ambiguous", () => {
  const dupes = [...CARDS, { clientid: "z", clientname: "Cooling Real Estate" }];
  assert.equal(parsePick('{"card":"Cooling Real Estate"}', dupes), null);
});

test("buildPickInput carries the sender, the hint, every card and the word json", () => {
  const input = buildPickInput({
    extractedName: "Cooling Pty Ltd", from: "x@console.com.au", fromName: "Cooling Real Estate",
    subject: "WO", address: "1 A St", propertyCards: ["Cooling Real Estate"], emailText: "a\n\n  b",
  }, CARDS);
  assert.match(input, /Sender: Cooling Real Estate <x@console\.com\.au>/);
  assert.match(input, /hint only\): Cooling Real Estate/);
  assert.match(input, /Email text:\na b/);
  for (const c of CARDS) assert.ok(input.includes(`\n${c.clientname}\n`));
  assert.match(input, /JSON/);
});

test("aiPickClient never throws and needs a client and a list", async () => {
  const failing = { responses: { create: async () => { throw new Error("401"); } } };
  assert.equal(await aiPickClient({ openai: failing, evidence: {}, cards: CARDS }), null);
  assert.equal(await aiPickClient({ openai: null, evidence: {}, cards: CARDS }), null);
  const ok = { responses: { create: async () => ({ output_text: '{"card":"Ray White ELG","reason":"signature"}' }) } };
  assert.equal(await aiPickClient({ openai: ok, evidence: {}, cards: [] }), null);
  assert.equal((await aiPickClient({ openai: ok, evidence: {}, cards: CARDS })).client.clientid, "b");
});

test("closestClients ranks on distinctive words, not 'Real Estate'", () => {
  assert.deepEqual(closestClients("Ray White", CARDS).map(c => c.clientid).sort(), ["a", "b"]);
  assert.deepEqual(closestClients("Morgan and Hayes Real Estate", CARDS).map(c => c.clientid), ["d"]);
  assert.deepEqual(closestClients("Jones Real Estate", CARDS), []);
});

import { test } from "node:test";
import assert from "node:assert/strict";
import { createOfficeSession, ensureTaskEmail, parseDirectEmailIds, findTaskIdByJobNumber } from "./aroflo-office.js";

process.env.AROFLO_V2_TOKEN = "t";
const env = { AROFLO_OFFICE_USER: "u", AROFLO_OFFICE_PASS: "p" };

const LOGIN_PAGE = '<form ref="loginForm" class="af-form">';
const TASK_PAGE  = '<div id="direct-email-container" data-zonename="TASKS" data-zonetype="47" data-zonesubtype="2" data-idcoded="92%26%3EAB%0A" data-orgid="92%26%3EOR%0A" data-showdialog="true"></div>';
const ADDRESS    = "baraelectrical_task+abc123@inboundemail.aroflo.com";

// A scripted fetch: each call pops the next reply; the log records what was asked.
function scripted(replies) {
  const log = [];
  const f = async (url, init = {}) => {
    log.push({ url: String(url), method: init.method ?? "GET", cookie: init.headers?.Cookie ?? "", body: init.body });
    const r = replies.shift();
    if (!r) throw new Error(`unexpected fetch ${url}`);
    const headers = new Headers(r.headers ?? {});
    return {
      ok: r.status < 300, status: r.status, headers: Object.assign(headers, { getSetCookie: () => r.setCookie ?? [] }),
      text: async () => r.body ?? "", json: async () => JSON.parse(r.body),
    };
  };
  return { f, log };
}

test("parseDirectEmailIds decodes both coded ids", () => {
  assert.deepEqual(parseDirectEmailIds(TASK_PAGE), { idcoded: "92&>AB\n", orgid: "92&>OR\n" });
  assert.equal(parseDirectEmailIds("<html>no control</html>"), null);
});

test("ensureTaskEmail returns an existing address without touching the office", async () => {
  const { f, log } = scripted([{ status: 200, body: JSON.stringify({ taskEmailAddress: ADDRESS }) }]);
  const session = createOfficeSession({ fetchImpl: f, env });
  assert.equal(await ensureTaskEmail("C8X1", session, f), ADDRESS);
  assert.equal(log.length, 1);
  assert.equal(session.hasSession(), false);
});

test("ensureTaskEmail logs in, reads the page, and posts the mint RPC", async () => {
  const { f, log } = scripted([
    { status: 200, body: JSON.stringify({ taskEmailAddress: null, _links: { office: "/ims/task?wrCoded=X" } }) },
    { status: 200, body: LOGIN_PAGE },                                                         // no cookie yet -> bounced
    { status: 200, body: JSON.stringify({ result: "success", message: "goto", data: "/ims/home" }), setCookie: ["CFID=1; Path=/", "CFTOKEN=2; Path=/"] },
    { status: 200, body: "<html>home</html>" },                                                // post-login goto
    { status: 200, body: TASK_PAGE },                                                          // retry with cookies
    { status: 200, body: JSON.stringify({ success: { value: true }, emailaddress: ADDRESS }) },
  ]);
  const session = createOfficeSession({ fetchImpl: f, env });
  assert.equal(await ensureTaskEmail("C8X1", session, f), ADDRESS);

  const loginCall = log[2];
  assert.match(loginCall.url, /\/ims\/Login\/login\.cfm$/);
  assert.deepEqual(JSON.parse(loginCall.body), { loginflag: 2, hostref: "office.aroflo.com", recaptcha: "", username: "u", password: "p", loginto: "office" });

  assert.equal(log[4].cookie, "CFID=1; CFTOKEN=2");
  const rpc = log[5];
  assert.match(rpc.url, /ZoneDirectEmailRpcController\.cfc\?method=save$/);
  assert.equal(rpc.body.get("targetid"), "92&>AB\n");
  assert.equal(rpc.body.get("orgid"), "92&>OR\n");
  assert.equal(rpc.body.get("zonetype"), "47");
  assert.equal(rpc.body.get("active"), "1");
});

test("ensureTaskEmail refuses to loop when the login itself fails", async () => {
  const { f } = scripted([
    { status: 200, body: JSON.stringify({ taskEmailAddress: null, _links: { office: "/ims/task" } }) },
    { status: 200, body: LOGIN_PAGE },
    { status: 200, body: JSON.stringify({ result: "error", message: "Too many login attempts. Please try again later." }) },
  ]);
  const session = createOfficeSession({ fetchImpl: f, env });
  await assert.rejects(ensureTaskEmail("C8X1", session, f), /Too many login attempts/);
});

test("findTaskIdByJobNumber finds a fresh job on page 1 and a typo nowhere", async () => {
  const items = [{ id: "C8X9", jobNumber: 107610 }, { id: "C8X8", jobNumber: 107609 }];
  const { f, log } = scripted([{ status: 200, body: JSON.stringify({ items }) }, { status: 200, body: JSON.stringify({ items }) }]);
  assert.equal(await findTaskIdByJobNumber("107609", f), "C8X8");
  assert.equal(await findTaskIdByJobNumber("999999", f), null);
  assert.equal(log.length, 2);
});

test("totp matches the RFC 6238 SHA-1 test vector", async () => {
  const { totp } = await import("./aroflo-office.js");
  // Secret "12345678901234567890" (base32 GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ) at T=59s -> 287082
  assert.equal(totp("GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ", 59_000), "287082");
  assert.equal(totp("GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ", 1_111_111_109_000), "081804");
});

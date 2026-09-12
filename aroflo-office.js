// AroFlo v2 REST helpers, plus the one thing v2 cannot do: give a task an inbound email
// address. AroFlo mints that address only when someone presses the "+" beside "Task Email
// Address" in the office UI, so this module logs into office.aroflo.com as a service user
// and presses it over HTTP. The chain, for a task the v1 API just created:
//
//   findTaskIdByJobNumber  ->  v2 task id (v1 and v2 ids are different id spaces)
//   ensureTaskEmail        ->  GET /v2/tasks/{id}; if taskEmailAddress is set, done
//                              else GET the task's office page (cookie) and read the two
//                              "coded" ids only that page carries, then POST the same RPC
//                              the button posts; its JSON reply carries the address
//
// Everything takes `fetch` so the tests can drive it without a network.

import { createHmac } from "node:crypto";

const V2_BASE     = "https://api.aroflo.com/v2";
const OFFICE_BASE = "https://office.aroflo.com";

// RFC 6238 TOTP, the formula every authenticator app runs: HMAC-SHA1 of the 30-second
// counter, dynamically truncated to 6 digits. AroFlo forces MFA on the service user, so the
// service is its own authenticator; the enrolled secret lives in AROFLO_OFFICE_TOTP_SECRET.
export function totp(base32Secret, now = Date.now()) {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let bits = 0, value = 0; const key = [];
  for (const ch of base32Secret.toUpperCase().replace(/=+$/, "")) {
    value = (value << 5) | alphabet.indexOf(ch); bits += 5;
    if (bits >= 8) { key.push((value >>> (bits - 8)) & 0xff); bits -= 8; }
  }
  const counter = Buffer.alloc(8);
  counter.writeBigUInt64BE(BigInt(Math.floor(now / 30000)));
  const mac = createHmac("sha1", Buffer.from(key)).update(counter).digest();
  const offset = mac[mac.length - 1] & 0x0f;
  const code = (mac.readUInt32BE(offset) & 0x7fffffff) % 1_000_000;
  return String(code).padStart(6, "0");
}

// ── v2 REST ─────────────────────────────────────────────────────────────────────────

function v2Headers() {
  const token = process.env.AROFLO_V2_TOKEN;
  if (!token) throw new Error("AROFLO_V2_TOKEN must be set");
  return { Authorization: `Bearer ${token}`, Accept: "application/json" };
}

export async function getTaskV2(v2TaskId, fetchImpl = fetch) {
  const res = await fetchImpl(`${V2_BASE}/tasks/${encodeURIComponent(v2TaskId)}`, { headers: v2Headers() });
  if (!res.ok) throw new Error(`AroFlo v2 GET task ${v2TaskId} failed: ${res.status} ${await res.text()}`);
  return res.json();
}

const JOB_NUMBER_PAGE_SIZE = 100;

// GET /v2/tasks has no job-number filter, but it sorts by job number, so page 1 descending
// starts at the highest number in the system and the page holding `jobNumber` is arithmetic
// from there. A job created moments ago is on page 1 — the common case here.
export async function findTaskIdByJobNumber(jobNumber, fetchImpl = fetch) {
  const target = Number(jobNumber);
  if (!Number.isInteger(target) || target <= 0) return null;

  const page = async (n) => {
    const res = await fetchImpl(`${V2_BASE}/tasks?sortBy=jobnumber&ascending=false&limit=${JOB_NUMBER_PAGE_SIZE}&page=${n}`, { headers: v2Headers() });
    if (!res.ok) throw new Error(`AroFlo v2 GET tasks page ${n} failed: ${res.status} ${await res.text()}`);
    return (await res.json()).items ?? [];
  };

  const first = await page(1);
  if (first.length === 0) return null;
  const onFirst = first.find(t => t.jobNumber === target);
  if (onFirst) return onFirst.id;
  if (target > first[0].jobNumber) return null;

  const guess = Math.floor((first[0].jobNumber - target) / JOB_NUMBER_PAGE_SIZE) + 1;
  for (const n of [guess, guess + 1, guess - 1].filter(n => n > 1)) {
    const hit = (await page(n)).find(t => t.jobNumber === target);
    if (hit) return hit.id;
  }
  return null;
}

// Attaches one file to the task's Documents & Photos. `bytes` is a Uint8Array/Buffer; the API
// wants raw base64 with no data: prefix. SHOW_ALL matches the visibility every photo the
// techs upload already has.
export async function uploadTaskDocument(v2TaskId, { filename, bytes, comment }, fetchImpl = fetch) {
  const res = await fetchImpl(`${V2_BASE}/tasks/${encodeURIComponent(v2TaskId)}/documents/upload`, {
    method: "POST",
    headers: { ...v2Headers(), "Content-Type": "application/json" },
    body: JSON.stringify({
      filename,
      base64: Buffer.from(bytes).toString("base64"),
      comment,
      documentFilter: "SHOW_ALL",
      useImageUploadResolution: true,
      imageResolution: 1600,
    }),
  });
  if (res.status !== 201) throw new Error(`AroFlo v2 document upload failed: ${res.status} ${await res.text()}`);
}

// ── office.aroflo.com session ───────────────────────────────────────────────────────

// One session per process. AroFlo suspends a login after repeated attempts, so this logs
// in once and only again when a request comes back bounced to the login page.
export function createOfficeSession({ fetchImpl = fetch, env = process.env } = {}) {
  const jar = new Map();

  function absorb(res) {
    for (const line of res.headers.getSetCookie?.() ?? []) {
      const [pair] = line.split(";");
      const eq = pair.indexOf("=");
      if (eq > 0) jar.set(pair.slice(0, eq).trim(), pair.slice(eq + 1).trim());
    }
  }
  const cookieHeader = () => [...jar].map(([k, v]) => `${k}=${v}`).join("; ");

  async function login() {
    const user = env.AROFLO_OFFICE_USER, pass = env.AROFLO_OFFICE_PASS;
    if (!user || !pass) throw new Error("AROFLO_OFFICE_USER and AROFLO_OFFICE_PASS must be set");
    jar.clear();
    // The login page is a Vue form; its XHR sends the JSON as a bare string, so match that.
    const res = await fetchImpl(`${OFFICE_BASE}/ims/Login/login.cfm`, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify({ loginflag: 2, hostref: "office.aroflo.com", recaptcha: "", username: user, password: pass, loginto: "office" }),
      redirect: "manual",
    });
    absorb(res);
    const body = await res.json().catch(() => null);
    if (!body || body.result !== "success") {
      throw new Error(`AroFlo office login failed: ${body?.message ?? `HTTP ${res.status}`}`);
    }
    // The reply names a post-login URL. That page is a chain of "complimentary" steps —
    // MFA, session limit, EULA, password change — each rendered as <form id="frmPostLogin">;
    // the session is only usable once it stops appearing.
    let page = body.message === "goto" && typeof body.data === "string" ? await officeGet(body.data, { retry: false }) : "";
    for (let step = 0; page.includes('id="frmPostLogin"'); step++) {
      if (step >= 4) throw new Error("AroFlo office login did not settle after 4 post-login steps");
      if (page.includes("<af-login-mfa-main-component")) {
        if (!env.AROFLO_OFFICE_TOTP_SECRET) throw new Error("AroFlo asks for MFA but AROFLO_OFFICE_TOTP_SECRET is not set — run enrol-authenticator.mjs");
        await verifyMfa({ mfatype: "GOOGLE_AUTHENTICATOR", token: totp(env.AROFLO_OFFICE_TOTP_SECRET) });
        page = await officeGet("/ims/PostLogin/html_post_login_home.cfm", { retry: false });
      } else if (page.includes("User Session Limit")) {
        // Sessions left behind by earlier process lifetimes; the form's Continue button ends them.
        page = await officePostForm("/ims/PostLogin/html_post_login_home.cfm", { ...hiddenFields(page), doTerminateSessions: "1" });
      } else {
        throw new Error("AroFlo office login is stuck on a post-login page this service cannot complete (EULA / password change) — log in as the service user in a browser once");
      }
    }
  }

  function hiddenFields(page) {
    const fields = {};
    for (const m of page.matchAll(/<input[^>]*type="hidden"[^>]*name="([^"]+)"[^>]*value="([^"]*)"/g)) fields[m[1]] = m[2];
    return fields;
  }

  // Same RPC the MFA screen calls. With `secret` it enrols the authenticator; without, it
  // checks a code against the enrolled one. trustdevice is asked for so a session that
  // outlives the process (unlikely, but free) needs no second code.
  async function verifyMfa(args) {
    const q = encodeURIComponent(JSON.stringify({ trustdevice: true, ...args }));
    const text = await officeGet(`/ims/rpc/mfa/MfaRpcController.cfc?method=verifyToken&argumentCollection=${q}`, { retry: false });
    let reply;
    try { reply = JSON.parse(text); } catch { throw new Error(`AroFlo MFA verify returned non-JSON: ${text.slice(0, 200)}`); }
    if (reply.success !== true) throw new Error(`AroFlo MFA verify failed: ${reply.message ?? text.slice(0, 200)}`);
    return reply;
  }

  // One-off: enrol an authenticator on the service user and hand the secret to `store`
  // (never printed). Run under `railway run` so store can push it into the service's vars.
  async function enrolAuthenticator(store) {
    await login();
    const gen = JSON.parse(await officeGet("/ims/rpc/mfa/MfaRpcController.cfc?method=generateSecret&argumentCollection=%7B%7D", { retry: false }));
    if (!gen.secret) throw new Error("AroFlo generateSecret returned no secret");
    await verifyMfa({ mfatype: "GOOGLE_AUTHENTICATOR", secret: gen.secret, token: totp(gen.secret) });
    await store(gen.secret);
  }

  function bouncedToLogin(res, text) {
    const loc = res.headers.get("location") ?? "";
    return res.status === 401 || /\/ims\/Login/i.test(loc) || (res.status === 200 && /ref="loginForm"/.test(text));
  }

  async function officeGet(path, { retry = true } = {}) {
    const url = path.startsWith("http") ? path : OFFICE_BASE + path;
    let res = await fetchImpl(url, { headers: { Cookie: cookieHeader() }, redirect: "manual" });
    absorb(res);
    // Follow same-host redirects by hand so every hop's cookies land in the jar.
    for (let hops = 0; res.status >= 300 && res.status < 400 && hops < 5; hops++) {
      const next = new URL(res.headers.get("location"), url).toString();
      if (/\/ims\/Login/i.test(next)) break;
      res = await fetchImpl(next, { headers: { Cookie: cookieHeader() }, redirect: "manual" });
      absorb(res);
    }
    const text = await res.text();
    if (bouncedToLogin(res, text)) {
      if (!retry) throw new Error("AroFlo office session not established");
      await login();
      return officeGet(path, { retry: false });
    }
    if (!res.ok) throw new Error(`AroFlo office GET ${path} failed: ${res.status}`);
    return text;
  }

  async function officePostForm(path, fields) {
    const form = new FormData();
    for (const [k, v] of Object.entries(fields)) form.append(k, v);
    const res = await fetchImpl(OFFICE_BASE + path, { method: "POST", headers: { Cookie: cookieHeader() }, body: form, redirect: "manual" });
    absorb(res);
    // A form post that redirects (the session-limit Continue does) hands back the page it lands on.
    if (res.status >= 300 && res.status < 400) return officeGet(new URL(res.headers.get("location"), OFFICE_BASE).toString(), { retry: false });
    const text = await res.text();
    if (bouncedToLogin(res, text)) throw new Error("AroFlo office session expired mid-request");
    if (!res.ok) throw new Error(`AroFlo office POST ${path} failed: ${res.status}`);
    return text;
  }

  return { login, officeGet, officePostForm, enrolAuthenticator, hasSession: () => jar.size > 0 };
}

// The office task page marks the direct-email control with the two ids the RPC wants.
// They are AroFlo's own "coded" ids — not the v1 or v2 API ids — and appear nowhere else.
export function parseDirectEmailIds(html) {
  const m = html.match(/id="direct-email-container"[^>]*?data-idcoded="([^"]*)"[^>]*?data-orgid="([^"]*)"/s);
  if (!m) return null;
  return { idcoded: decodeURIComponent(m[1]), orgid: decodeURIComponent(m[2]) };
}

// Returns the task's inbound email address, minting one if the task has none yet.
export async function ensureTaskEmail(v2TaskId, session, fetchImpl = fetch) {
  const task = await getTaskV2(v2TaskId, fetchImpl);
  if (task.taskEmailAddress) return task.taskEmailAddress;

  const officePath = task._links?.office;
  if (!officePath) throw new Error(`v2 task ${v2TaskId} has no _links.office`);

  const page = await session.officeGet(officePath);
  // A user who has never logged in through a browser is parked on the EULA page instead.
  if (page.includes('id="frmPostLogin"')) throw new Error("AroFlo office login is stuck on a first-login page (EULA / password change) — log in as the service user in a browser once and accept it");
  const ids = parseDirectEmailIds(page);
  if (!ids) throw new Error(`Office page for task ${v2TaskId} has no direct-email control (permission, or AroFlo changed the page)`);

  // Same payload the "+" button sends: zonetype 47 / subtype 2 is the Tasks zone.
  const reply = await session.officePostForm("/ims/rpc/zonedirectemail/ZoneDirectEmailRpcController.cfc?method=save", {
    targetid: ids.idcoded, orgid: ids.orgid, zonetype: "47", zonesubtype: "2",
    fromfilter: "", sendeventmessage: "1", active: "1",
  });
  let parsed;
  try { parsed = JSON.parse(reply); } catch { throw new Error(`Direct-email RPC returned non-JSON: ${reply.slice(0, 200)}`); }
  if (!parsed.emailaddress) throw new Error(`Direct-email RPC did not return an address: ${reply.slice(0, 200)}`);
  return parsed.emailaddress;
}

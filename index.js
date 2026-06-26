import "dotenv/config";
import express from "express";
import OpenAI from "openai";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { createHmac } from "crypto";
import { PACKAGE_TEMPLATES } from "./templates.js";

const REQUIRED_ENV = [
  "OPENAI_API_KEY",
  "GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET", "GRAPH_RECIPIENT",
  "UENCODED", "PENCODED", "ORGENCODED", "SECRET_KEY",
];
for (const key of REQUIRED_ENV) {
  if (!process.env[key]) {
    console.error(`${key} is not set`);
    process.exit(1);
  }
}

const TRIGGER_CATEGORY          = "Bara AI";
const PROCESSING_CATEGORY       = "Processing";
const CLIENT_NOT_FOUND_CATEGORY = "Client not found";
const READING_EMAIL_CATEGORY    = "Reading Email";
const SENDING_TO_AI_CATEGORY    = "Sending to AI";
const CREATING_JOB_CATEGORY     = "Creating Job";
const RICA_CATEGORY             = "Rica";
const WORKORDERS_EMAIL          = "workorders@baraelectrical.com.au";
const BRANDON_EMAIL             = "brandon.roberts@baraelectrical.com.au";
const POLL_INTERVAL_MS          = 5 * 60 * 1000;

// All transient status categories — stripped when transitioning to next stage
const STATUS_CATEGORIES = [
  PROCESSING_CATEGORY, READING_EMAIL_CATEGORY,
  SENDING_TO_AI_CATEGORY, CREATING_JOB_CATEGORY,
  CLIENT_NOT_FOUND_CATEGORY,
];

let pollRunning = false;

const app = express();
app.use(express.json({ limit: "25mb" }));

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// ================================================================
// GRAPH API
// ================================================================
let tokenCache = { token: null, expiry: 0 };

async function getAccessToken() {
  if (tokenCache.token && Date.now() < tokenCache.expiry - 60000) return tokenCache.token;
  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     process.env.GRAPH_CLIENT_ID,
        client_secret: process.env.GRAPH_CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default",
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error(`Graph auth failed: ${JSON.stringify(data)}`);
  tokenCache = { token: data.access_token, expiry: Date.now() + data.expires_in * 1000 };
  return tokenCache.token;
}

async function graphFetch(path, options = {}) {
  const token = await getAccessToken();
  return fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization:  `Bearer ${token}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });
}

// ================================================================
// AROFLO API
// ================================================================
const AROFLO_BASE   = "https://api.aroflo.com/";
const AROFLO_ACCEPT = "text/json";

function arofloAuth() {
  return (
    "uencoded="   + encodeURIComponent(process.env.UENCODED) +
    "&pencoded="  + encodeURIComponent(process.env.PENCODED) +
    "&orgEncoded=" + encodeURIComponent(process.env.ORGENCODED)
  );
}

function arofloSign(method, query, ts) {
  return createHmac("sha512", process.env.SECRET_KEY)
    .update([method, "", AROFLO_ACCEPT, arofloAuth(), ts, query].join("+"))
    .digest("hex");
}

async function arofloGet(params) {
  const ts   = new Date().toISOString();
  const auth = arofloAuth();
  const res  = await fetch(AROFLO_BASE + "?" + params, {
    headers: {
      Accept:          AROFLO_ACCEPT,
      Authorization:   auth,
      Authentication:  "HMAC " + arofloSign("GET", params, ts),
      afdatetimeutc:   ts,
    },
  });
  const data = await res.json();
  if (data.status !== "0") throw new Error(`Aroflo GET failed: ${data.statusmessage}`);
  return data.zoneresponse;
}

async function arofloPost(body) {
  const ts   = new Date().toISOString();
  const auth = arofloAuth();
  const res  = await fetch(AROFLO_BASE + "?", {
    method: "POST",
    headers: {
      Accept:          AROFLO_ACCEPT,
      Authorization:   auth,
      Authentication:  "HMAC " + arofloSign("POST", body, ts),
      afdatetimeutc:   ts,
      "Content-Type":  "application/x-www-form-urlencoded",
    },
    body,
  });
  const data = await res.json();
  if (data.status !== "0") throw new Error(`Aroflo POST failed: ${data.statusmessage}`);
  return data.zoneresponse;
}

// Task type IDs sourced from live Aroflo system on 2026-06-24
const TASK_TYPE_MAP = {
  "EC1":                            "JCYqLyBSQCAgCg==", // $120 Standard Electrical Compliance
  "AC1":                            "JCYqWyVQUCAgCg==", // $180 Standard Air Con Service
  "AC2":                            "JCYqLyNRMCAgCg==", // $270 Deluxe Air Con Service
  "ACEC1":                          "JCYqLyBRMCAgCg==", // $240 Air Con & Electrical Compliance
  "Real Estate Aircon Maintenance":  "JCZaSyBSQCAgCg==", // Real-Estate Air Con Maintenance
  "Real Estate General Maintenance": "JCYqWyVQICAgCg==", // Real-Estate General Maintenance
};

const SUBSTATUS_MAP = {
  "EC1":                            "IyQ6SyYK", // Ready to schedule PPM1
  "AC1":                            "IyQ6SyYK", // Ready to schedule PPM1
  "AC2":                            "Iyc6LyUK", // Ready to schedule (Specialised)
  "ACEC1":                          "IyQ6SyYK", // Ready to schedule PPM1
  "Real Estate Aircon Maintenance":  "Iyc6LyYK", // Ready to schedule
  "Real Estate General Maintenance": "Iyc6LyYK", // Ready to schedule
};

// Fetch contacts for a client using join=contacts, then match by PM name.
async function findContact(clientId, pmName) {
  if (!clientId || !pmName) return null;
  try {
    const zone = await arofloGet(
      "zone=clients" +
      "&join=" + encodeURIComponent("contacts") +
      "&where=" + encodeURIComponent(`and|clientid|=|${clientId}`) +
      "&where=" + encodeURIComponent("and|archived|=|false") +
      "&page=1"
    );
    const raw     = zone.clients;
    const client  = Array.isArray(raw) ? raw[0] : raw;
    const contacts = client?.contacts || [];
    const arr      = Array.isArray(contacts) ? contacts : [contacts];
    const nameLower = pmName.toLowerCase();
    return arr.find(c => `${c.givennames} ${c.surname}`.toLowerCase().includes(nameLower)) || null;
  } catch (err) {
    console.warn("Contact search failed:", err.message);
    return null;
  }
}

// Search Aroflo for a client by name using the API where clause.
// Tries progressively shorter variants: full name, before-pipe, first word.
async function findClient(realEstateName) {
  if (!realEstateName) return null;

  const candidates = [
    realEstateName,                              // "Realmark Urban"
    realEstateName.split(/[|,]/)[0].trim(),      // before " | " separator
    realEstateName.split(" ")[0],                // first word only
  ].filter((v, i, a) => a.indexOf(v) === i);    // dedupe

  for (const name of candidates) {
    const zone    = await arofloGet(
      "zone=clients&where=" + encodeURIComponent(`and|clientname|=|${name}`) + "&page=1"
    );
    const raw     = zone.clients;
    if (!raw) continue;
    const arr     = Array.isArray(raw) ? raw : [raw];
    if (arr.length > 0) return arr[0];
  }

  return null;
}

// Find a location by street address under a client, then update SiteContact/SitePhone if stale.
// Fetches all locations linked to the client and matches by street address locally —
// the locationname|like WHERE clause is not supported by Aroflo.
async function findOrUpdateLocation(clientId, address, tenantName, tenantContact) {
  if (!address) return null;

  // Strip unit prefix: "1412/380 Murray Street, Perth WA" → "380 Murray Street"
  const streetPart = address.replace(/^\d+\//, "").split(",")[0].trim().toLowerCase();

  let zone;
  try {
    zone = await arofloGet(
      "zone=locations" +
      "&where=" + encodeURIComponent(`and|linkedtoid|=|${clientId}`) +
      "&page=1"
    );
  } catch (err) {
    console.warn("Location search failed:", err.message);
    return null;
  }

  const raw = zone.locations;
  if (!raw) { console.log("No locations found for client"); return null; }

  const all      = Array.isArray(raw) ? raw : [raw];
  const location = all.find(l => l.locationname?.toLowerCase().includes(streetPart));

  if (!location) {
    console.log("No location matching:", streetPart, "— available:", all.map(l => l.locationname));
    return null;
  }

  console.log("FOUND LOCATION:", location.locationid, location.locationname);

  const needsUpdate =
    (tenantName    && location.SiteContact !== tenantName) ||
    (tenantContact && location.SitePhone   !== tenantContact);

  if (needsUpdate) {
    console.log("UPDATING TENANT — was:", location.SiteContact, "/", location.SitePhone);
    const xml =
`<locations>
  <location>
    <locationid>${location.locationid}</locationid>
    ${tenantName    ? `<SiteContact>${tenantName}</SiteContact>`  : ""}
    ${tenantContact ? `<SitePhone>${tenantContact}</SitePhone>`   : ""}
  </location>
</locations>`;
    try {
      await arofloPost("zone=locations&postxml=" + encodeURIComponent(xml));
      console.log("TENANT DETAILS UPDATED");
    } catch (err) {
      console.warn("Tenant update failed:", err.message);
    }
  }

  return location;
}

function buildDescription(result) {
  const parts = [];
  const spacer = `<p>&nbsp;</p>`;

  if (result["task-description"] || result["task-type"]) {
    parts.push(`<p>${result["task-description"] || result["task-type"]}</p>`);
  }

  const hasHighlights = result["expenditure-limit"] || result["access-details"];
  if (hasHighlights) parts.push(spacer);

  if (result["expenditure-limit"]) {
    parts.push(`<p><span style="background:#cce5ff;font-weight:bold">Expenditure Limit: ${result["expenditure-limit"]}</span></p>`);
  }

  if (result["access-details"]) {
    parts.push(`<p><span style="background:#ccffcc;font-weight:bold">Access Details: ${result["access-details"]}</span></p>`);
  }

  // Fall back to task-type if AI forgot to set package (e.g. task-type is EC1 but package is null)
  const pkg = (result["package"] && result["package"] !== "null" ? result["package"] : null)
    ?? (PACKAGE_TEMPLATES[result["task-type"]] ? result["task-type"] : null);
  if (pkg && PACKAGE_TEMPLATES[pkg]) {
    parts.push(spacer);
    parts.push(PACKAGE_TEMPLATES[pkg]);
  }

  return parts.join("\n");
}

async function createArofloJob(result, rawEmail) {
  console.log("CREATING AROFLO JOB...");

  const taskTypeId  = TASK_TYPE_MAP[result["task-type"]];
  const substatusId = SUBSTATUS_MAP[result["task-type"]] || "Iyc6LyYK"; // default: Ready to schedule
  if (!taskTypeId) console.warn("Unknown task type:", result["task-type"]);

  const client = await findClient(result["real-estate"]);
  if (!client) throw new Error(`Client not found in Aroflo: "${result["real-estate"]}"`);
  console.log("CLIENT:", client.clientid, client.clientname);

  const location = await findOrUpdateLocation(
    client.clientid,
    result.address,
    result["tenant-name"],
    result["tenant-contact"]
  );

  const pmContact = await findContact(client.clientid, result["property-manager"]);
  if (pmContact) console.log("PM CONTACT:", pmContact.contactid, pmContact.contactname);
  else console.warn("PM contact not found in Aroflo:", result["property-manager"]);

  const notes = [
    result["order-number"]     ? `Work Order: ${result["order-number"]}`          : null,
    result["tenant-name"]      ? `Tenant: ${result["tenant-name"]}`                : null,
    result["tenant-contact"]   ? `Tenant Contact: ${result["tenant-contact"]}`     : null,
    result["access-details"]      ? `Access Details: ${result["access-details"]}`           : null,
    result["expenditure-limit"]   ? `Expenditure Limit: ${result["expenditure-limit"]}`     : null,
    result["property-manager"]    ? `Property Manager: ${result["property-manager"]}`       : null,
  ].filter(Boolean).join("\n");

  const dueDate = (() => {
    const d = new Date();
    d.setDate(d.getDate() + 30);
    return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
  })();

  // "36 Dollis Way, Kingsley, WA 6026" → "36 Dollis Way Kingsley"
  const taskName = result.address
    ? result.address.split(",").slice(0, 2).map(p => p.trim()).join(" ")
    : "";

  const xml =
`<tasks>
  <task>
    <org><orgid>JiYqTydSXDcmCg==</orgid></org>
    <contactname>Bara AI</contactname>
    ${taskTypeId ? `<tasktype><tasktypeid>${taskTypeId}</tasktypeid></tasktype>`                  : ""}
    <client><clientid>${client.clientid}</clientid></client>
    ${pmContact ? `<contact><userid>${pmContact.userid}</userid></contact>` : ""}
    ${location ? `<location><locationid>${location.locationid}</locationid></location>` : ""}
    ${result.address && !location  ? `<sitename>${result.address}</sitename>`          : ""}
    <taskname>${taskName}</taskname>
    <description><![CDATA[${buildDescription(result)}]]></description>
    <duedate>${dueDate}</duedate>
    ${result["order-number"] ? `<custon>${result["order-number"]}</custon>` : ""}
    ${(result["account-to"] || result["real-estate"]) ? `<customfields><customfield><name><![CDATA[ Account To: ]]></name><type><![CDATA[ text ]]></type><value><![CDATA[${result["account-to"] || result["real-estate"]}]]></value></customfield></customfields>` : ""}
  </task>
</tasks>`;

  const zone     = await arofloPost("zone=tasks&postxml=" + encodeURIComponent(xml));
  const pr          = zone.postresults;
  const insertTotal = Number(pr?.inserttotal ?? 0);

  if (insertTotal < 1) {
    const errors = pr?.errors;
    const msgs   = errors && (Array.isArray(errors) ? errors : [errors]).length
      ? (Array.isArray(errors) ? errors : [errors]).map(e => e.detail || e.message || JSON.stringify(e)).join("; ")
      : "No job inserted";
    throw new Error(`Aroflo task creation failed: ${msgs}`);
  }

  const inserted = pr?.inserts?.tasks;
  const task     = Array.isArray(inserted) ? inserted[0] : inserted;
  const taskId   = task?.taskid;

  // jobnumber isn't in the insert response — fetch it. Also use the fetched taskId
  // for note posting since it's the same format /test-note uses and is confirmed working.
  let jobNumber = "(see Aroflo)";
  let confirmedTaskId = taskId;
  if (taskId) {
    try {
      const fetched = await arofloGet("zone=tasks&where=" + encodeURIComponent(`and|taskid|=|${taskId}`) + "&page=1");
      const arr = Array.isArray(fetched.tasks) ? fetched.tasks : [fetched.tasks];
      jobNumber = arr[0]?.jobnumber || taskId;
      confirmedTaskId = arr[0]?.taskid || taskId;
    } catch (err) {
      console.warn("Could not fetch job number:", err.message);
      jobNumber = taskId;
    }
  }
  console.log("AROFLO JOB CREATED — job number:", jobNumber, "taskId:", confirmedTaskId);

  // Post the original email as a note on the task
  if (confirmedTaskId && rawEmail) {
    try {
      const noteHtml = await emailHtmlForNote(rawEmail);
      const xml = `<tasks><task><taskid>${confirmedTaskId}</taskid><notes><note><content><![CDATA[${noteHtml}]]></content></note></notes></task></tasks>`;
      await arofloPost("zone=tasks&postxml=" + encodeURIComponent(xml));
      console.log("Email note posted");
    } catch (err) {
      console.warn("Email note post failed:", err.message);
    }
  }

  // Aroflo doesn't apply substatus on create — do a follow-up update
  if (confirmedTaskId && substatusId) {
    const updateXml =
`<tasks>
  <task>
    <taskid>${confirmedTaskId}</taskid>
    <status>not started</status>
    <substatus><substatusid>${substatusId}</substatusid></substatus>
  </task>
</tasks>`;
    const upZone = await arofloPost("zone=tasks&postxml=" + encodeURIComponent(updateXml));
    const upPr   = upZone.postresults;
    if (Number(upPr?.updatetotal ?? 0) > 0) {
      console.log("Substatus set:", substatusId);
    } else {
      console.warn("Substatus update returned 0 updates — may not have applied");
    }
  }

  return { zone, jobNumber };
}

// ================================================================
// PDF + HTML helpers
// ================================================================
async function extractPDF(data) {
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf         = await loadingTask.promise;
  let text = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page    = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(item => item.str).join(" ") + "\n";
  }
  return text;
}

function cleanHtml(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Decode SafeLinks/Inky wrapped URLs. Inky requires following the redirect.
async function decodeWrappedLinks(html) {
  const matches = [...html.matchAll(/href="([^"]+)"/gi)];
  const unique  = [...new Set(matches.map(m => m[1]))];
  const map     = {};

  await Promise.all(unique.map(async href => {
    const decoded = href.replace(/&amp;/g, "&");
    try {
      if (/safelinks\.protection\.outlook\.com/i.test(decoded)) {
        const url = new URL(decoded).searchParams.get("url");
        if (url) map[href] = decodeURIComponent(url);
      } else if (/shared\.outlook\.inky\.com/i.test(decoded)) {
        const res = await fetch(decoded.includes("confirm=True") ? decoded : decoded + "&confirm=True", {
          redirect: "follow",
          headers: { "User-Agent": "Mozilla/5.0" },
        });
        map[href] = res.url;
      }
    } catch { /* leave as-is */ }
  }));

  return html.replace(/href="([^"]+)"/gi, (match, href) =>
    map[href] ? `href="${map[href]}"` : match
  );
}

// Strip scripts/styles/tracking pixels but keep HTML structure for display in Aroflo notes
async function emailHtmlForNote(html) {
  let cleaned = html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<img[^>]*>/gi, "")
    .trim();

  // Remove Inky security banner — Inky wraps its banner between x_ipw-middle-* and x_ipw-end-*
  // anchors. Strip everything from the start up to and including the end marker + closing div.
  cleaned = cleaned.replace(/[\s\S]*?<a\s[^>]*name="x_ipw-end-[^"]*"[^>]*><\/a><\/div>\s*/i, "");

  cleaned = await decodeWrappedLinks(cleaned);

  return `<p><strong>Work Order Email</strong></p>${cleaned}`;
}

// Known work order portal domains
const WORKORDER_DOMAINS = /tapihq\.com|propertytree\.com|propertyme\.com\.au|console\.net\.au|inspection\.express/i;

// Search raw HTML anchor tags — match on link text containing "work order" or "workorder".
function findWorkOrderLink(rawHtml) {
  const unescaped = rawHtml.replace(/&amp;/g, "&");

  // Extract <a href="...">text</a> pairs
  const anchors = [...unescaped.matchAll(/<a\s[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi)];

  for (const [, href, rawText] of anchors) {
    if (!href.startsWith("https://")) continue;
    const text = rawText.replace(/<[^>]*>/g, "").trim();
    if (/work\s*order/i.test(text)) {
      console.log("SELECTED WORKORDER LINK by text:", text.slice(0, 60), "->", href.slice(0, 80));
      return href;
    }
  }

  // Fallback: known portal domains or "workorder" in the URL itself
  for (const [, href] of anchors) {
    if (!href.startsWith("https://")) continue;
    const dest = /safelinks\.protection\.outlook\.com/i.test(href)
      ? (() => { try { return decodeURIComponent(new URL(href).searchParams.get("url") || href); } catch { return href; } })()
      : href;
    if (/workorder/i.test(dest) || WORKORDER_DOMAINS.test(dest)) {
      console.log("SELECTED WORKORDER LINK by domain/url:", dest.slice(0, 80));
      return href;
    }
  }

  // Last resort: plain-text TAPI URL
  const plainText = unescaped.replace(/<[^>]*>/g, " ");
  const m = plainText.match(/https:\/\/url\d+\.tapihq\.com\/ls\/click\S+/i);
  if (m) return m[0].split(/[>")\s]/)[0].trim();

  return null;
}

// ================================================================
// EMAIL PROCESSING
// ================================================================
async function processMessage(message, mailbox = process.env.GRAPH_RECIPIENT, onStatus = null) {
  console.log("WORK ORDER EMAIL RECEIVED:", message.subject);

  const attachments = message.attachments || [];
  const rawBody     = message.body?.content || "";
  const emailBodyText = (message.body?.contentType === "html" ? cleanHtml(rawBody) : rawBody).replace(/\s+/g, " ").trim();
  let   textForAI     = emailBodyText;

  function withEmailBody(primary) {
    return `--- WORK ORDER CONTENT (prefer this) ---\n${primary}\n\n--- EMAIL BODY (use for anything not found above) ---\n${emailBodyText}`;
  }

  // Work order link detection
  const workOrderLink = findWorkOrderLink(rawBody);

  if (workOrderLink) {
    console.log("WORK ORDER LINK FOUND:", workOrderLink.slice(0, 120));
    try {
      let response = await fetch(workOrderLink, {
        redirect: "follow",
        headers: { "User-Agent": "Mozilla/5.0" },
      });

      // Inky link protection stops the redirect — bypass with confirm=True
      if (response.url.includes("shared.outlook.inky.com") && !response.url.includes("confirm=True")) {
        console.log("INKY BYPASS:", response.url.slice(0, 120));
        response = await fetch(response.url + "&confirm=True", {
          redirect: "follow",
          headers: { "User-Agent": "Mozilla/5.0" },
        });
      }

      console.log("FINAL URL:", response.url);
      const contentType = response.headers.get("content-type") || "";
      if (contentType.includes("pdf")) {
        const buffer  = await response.arrayBuffer();
        const pdfText = await extractPDF(new Uint8Array(buffer));
        textForAI = withEmailBody(pdfText.replace(/\s+/g, " ").trim());
        console.log("USING LINKED PDF + EMAIL BODY FOR AI");
      } else {
        const html     = await response.text();
        const linkText = cleanHtml(html).slice(0, 50000);
        console.log("LINK CLEAN TEXT LENGTH:", linkText.length);
        if (linkText.length > 200) {
          textForAI = withEmailBody(linkText);
          console.log("USING LINK CONTENT + EMAIL BODY FOR AI");
        } else {
          console.log("LINK CONTENT TOO SHORT — FALLING BACK TO EMAIL BODY");
        }
      }
    } catch (err) {
      console.error("LINK FETCH ERROR:", err.message);
      console.log("FALLING BACK TO EMAIL BODY");
    }
  }

  // Work order PDF attachment — takes priority over link
  const workorderAttachment = attachments.find(a =>
    (a.name || "").toLowerCase().endsWith(".pdf")
  );

  if (workorderAttachment) {
    console.log("FOUND WORK ORDER PDF:", workorderAttachment.name);
    const attachRes  = await graphFetch(
      `/users/${mailbox}/messages/${message.id}/attachments/${workorderAttachment.id}`
    );
    const attachData = await attachRes.json();
    const data       = Uint8Array.from(atob(attachData.contentBytes), c => c.charCodeAt(0));
    const pdfText    = await extractPDF(data);
    textForAI = withEmailBody(pdfText.replace(/\s+/g, " ").trim());
    console.log("USING PDF ATTACHMENT + EMAIL BODY FOR AI");
  } else if (!workOrderLink) {
    console.log("NO PDF OR LINK - USING EMAIL BODY ONLY");
  }

  console.log("TEXT LENGTH:", textForAI.length);
  if (onStatus) await onStatus(SENDING_TO_AI_CATEGORY);

  const responseAI = await openai.responses.create({
    model: "gpt-5-mini",
    text: { format: { type: "json_object" } },
    instructions: `You are a work order extraction system for an electrical company in Australia.

CRITICAL RULES:
- tenant-name and tenant-contact come from the Tenant Details section OR any section labelled "Contact for job access" or similar. If no tenant is listed and the property is explicitly stated as vacant, set tenant-name to "Vacant" and ensure access-details captures any lockbox or key collection info. If no tenant is listed and the property is NOT stated as vacant, leave tenant-name null. Never use access details, lockbox info, or key numbers as the tenant name. If multiple tenants are listed, include ALL of them separated by commas — do not drop any.
- access-details is ONLY physical access codes/numbers — key numbers, lockbox codes, gate codes, swipe card numbers. e.g. "Key: 1234", "Lockbox code: 56", "Gate code: 789". Do NOT include contact instructions, tenant names, safety instructions, or anything that is not a physical code or number.
- expenditure-limit is the dollar amount only — e.g. "$330". Strip any conditions, notes, or extra text after the amount.
- confidence is a float 0.0–1.0 rating how confident you are in the overall extraction. 1.0 = all fields clearly present, 0.0 = guessing most fields.
- notes is any concerns, ambiguities, or flags worth mentioning — e.g. missing fields, conflicting info, unusual job details. Leave null if nothing to flag.
- tenant-contact must contain phone numbers ONLY — no names, no labels, just the numbers. If there are multiple, separate with commas. Prefer mobile over home numbers. Australian numbers always start with 0 (e.g. 0412 345 678) — always include the leading 0.
- property-manager comes from the Property Manager section, OR from an Agency Details section where the manager is listed (e.g. "Manager: Jane Smith"). Use the person's name only, not the agency name.
- account-to must include ALL owners exactly as written, always in the format: owners c/o real estate.
- real-estate must always be a company or agency name — never a URL or domain. If the source contains something like "aussieproperty.com.au", convert it to a readable name (e.g. "Aussie Property") by stripping the domain extension and formatting as a proper name. If you cannot find it directly, look for it in account-to after the c/o.
- order-number is the job/work order number.
- task-description must be a concise electrician job summary. If anything is listed as conditional or requires approval (e.g. "deluxe clean if approved", "AC2 if required"), include that in the description too.
- Do NOT guess missing fields — if missing return null.
- The text may be a structured form (with clear sections) OR plain prose in an email. Extract the same fields either way — don't return null just because sections aren't labelled.
- If the input contains both a PDF CONTENT section and an EMAIL BODY section, prefer the PDF for all fields but check the email body for anything not found in the PDF.
- For task-type: if the text explicitly mentions EC1, AC1, AC2, or ACEC1 anywhere, use that — it takes priority over everything else, even if other work is also mentioned. If a compliance check or aircon service is recommended or suggested (e.g. "recommend compliance check", "suggest an EC1"), treat it as that task type — recommendations from a PM or owner should be acted on. EXCEPTION: if the work order includes a package (EC1/AC1/AC2/ACEC1) AND additional work beyond the package, set task-type to "Real Estate General Maintenance" (or "Real Estate Aircon Maintenance" if the extra work is aircon-related) — but still set the package field to the package code.
- package: set to "EC1", "AC1", "AC2", or "ACEC1" if the job involves one of these standard packages — even if task-type has been set to general/aircon maintenance due to extra work being present. Only return null if there is genuinely no package involved at all (e.g. a pure repair or general maintenance job with no compliance check or aircon service).

TASK TYPES:
EC1 = Electrical Compliance Check only
AC1 = Aircon Servicing only
AC2 = Deluxe Aircon Clean
ACEC1 = Aircon service AND electrical compliance check combined
Real Estate Aircon Maintenance = aircon related jobs
Real Estate General Maintenance = everything else

Return ONLY valid JSON with these exact keys:
{
  "task-type": "",
  "tenant-name": "",
  "tenant-contact": "",
  "address": "",
  "task-description": "",
  "real-estate": "",
  "property-manager": "",
  "account-to": "",
  "order-number": "",
  "access-details": "",
  "expenditure-limit": "",
  "package": "EC1 | AC1 | AC2 | ACEC1 | null",
  "confidence": 0.0,
  "notes": ""
}`,
    input: `Extract the following work order and return JSON:\n\n${textForAI}`,
  });

  const parsed = JSON.parse(responseAI.output_text);

  // Normalise address to title case in case the source had ALL CAPS suburb
  if (parsed.address) {
    parsed.address = parsed.address.replace(/\b\w+/g, w =>
      w.length <= 2 ? w.toUpperCase() : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()
    );
    // Expand common street type abbreviations
    const streetAbbr = {
      "\\bSt\\b": "Street", "\\bRd\\b": "Road",    "\\bAve\\b": "Avenue",
      "\\bDr\\b": "Drive",  "\\bLn\\b": "Lane",     "\\bCt\\b": "Court",
      "\\bPl\\b": "Place",  "\\bCl\\b": "Close",    "\\bCres\\b": "Crescent",
      "\\bBlvd\\b": "Boulevard", "\\bHwy\\b": "Highway", "\\bFwy\\b": "Freeway",
      "\\bTce\\b": "Terrace",    "\\bPde\\b": "Parade",  "\\bGr\\b": "Grove",
      "\\bBvd\\b": "Boulevard",  "\\bCct\\b": "Circuit", "\\bEsp\\b": "Esplanade",
    };
    for (const [abbr, full] of Object.entries(streetAbbr)) {
      parsed.address = parsed.address.replace(new RegExp(abbr, "gi"), full);
    }
  }

  // If AI returned a URL/domain for real-estate, strip the TLD and convert to a readable name
  if (parsed["real-estate"] && /\.(com\.au|net\.au|org\.au|com|net|org|io|au)$/i.test(parsed["real-estate"])) {
    parsed["real-estate"] = parsed["real-estate"]
      .replace(/\.(com\.au|net\.au|org\.au|com|net|org|io|au)$/i, "")
      .replace(/[-_.]/g, " ")
      .replace(/\b\w/g, c => c.toUpperCase())
      .trim();
  }

  // Australian mobile numbers are 10 digits starting with 0 — if AI drops the leading 0, restore it
  if (parsed["tenant-contact"]) {
    parsed["tenant-contact"] = parsed["tenant-contact"]
      .split(",")
      .map(n => {
        const digits = n.replace(/\D/g, "");
        return digits.length === 9 && digits.startsWith("4") ? "0" + n.trim() : n.trim();
      })
      .join(", ");
  }

  return { result: parsed, rawEmail: rawBody };
}

// ================================================================
// RICA FORWARD LOOP
// ================================================================
async function forwardRicaEmails() {
  try {
    // Read Rica-tagged emails from workorders inbox — read-only, nothing changes there
    const filter = encodeURIComponent(`categories/any(c:c eq '${RICA_CATEGORY}')`);
    const res    = await graphFetch(
      `/users/${WORKORDERS_EMAIL}/mailFolders/inbox/messages` +
      `?$filter=${filter}&$select=id,subject,body,hasAttachments&$top=50`
    );
    const data = await res.json();
    if (!res.ok) throw new Error(`Graph API error ${res.status}: ${JSON.stringify(data?.error)}`);

    const emails = data.value || [];
    if (emails.length) console.log(`Rica: ${emails.length} tagged email(s) in workorders inbox`);

    const aiTestingFolder = await getAiTestingFolderId();
    const aiDoneFolder    = await getAiDoneFolderId();

    for (const email of emails) {
      const fwdSubject  = `AI testing - FW: ${email.subject}`;
      const safeSubj    = fwdSubject.replace(/'/g, "''");
      const dedupFilter = encodeURIComponent(`subject eq '${safeSubj}'`);

      // Dedup: check "AI testing" folder (unprocessed) and "AI done" folder (already processed)
      const testingFolder = aiTestingFolder
        ? `/users/${BRANDON_EMAIL}/mailFolders/${aiTestingFolder}/messages`
        : `/users/${BRANDON_EMAIL}/mailFolders/inbox/messages`;
      const dedupRes  = await graphFetch(`${testingFolder}?$filter=${dedupFilter}&$select=id&$top=1`);
      const dedupData = await dedupRes.json();
      if (dedupData.value?.length > 0) continue;

      if (aiDoneFolder) {
        const doneRes  = await graphFetch(`/users/${BRANDON_EMAIL}/mailFolders/${aiDoneFolder}/messages?$filter=${dedupFilter}&$select=id&$top=1`);
        const doneData = await doneRes.json();
        if (doneData.value?.length > 0) continue;
      }

      // Fetch attachments if present
      let attachments = [];
      if (email.hasAttachments) {
        const attRes  = await graphFetch(`/users/${WORKORDERS_EMAIL}/messages/${email.id}/attachments`);
        const attData = await attRes.json();
        attachments = (attData.value || []).map(a => ({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name:         a.name,
          contentType:  a.contentType,
          contentBytes: a.contentBytes,
        }));
      }

      // Strip Graph metadata from body before sending — passing email.body directly
      // can cause Exchange to ignore the custom subject
      const message = {
        subject:      fwdSubject,
        body:         { contentType: email.body?.contentType || "html", content: email.body?.content || "" },
        toRecipients: [{ emailAddress: { address: BRANDON_EMAIL } }],
      };
      if (attachments.length) message.attachments = attachments;

      console.log("Rica: sending with subject:", fwdSubject);
      const sendRes = await graphFetch(`/users/${WORKORDERS_EMAIL}/sendMail`, {
        method: "POST",
        body: JSON.stringify({ message, saveToSentItems: false }),
      });

      if (!sendRes.ok) {
        const sendErr = await sendRes.json().catch(() => ({}));
        console.warn("Rica send failed:", email.subject, sendRes.status, JSON.stringify(sendErr?.error));
        continue;
      }
      console.log("Rica sent to Brandon:", fwdSubject);
    }
  } catch (err) {
    console.error("Rica forward error:", err.message);
  }
}

// ================================================================
// AI TESTING LOOP — process emails in Brandon's "AI testing" folder,
// reply with extracted info, no Aroflo job created
// ================================================================
let aiTestingFolderId  = null;
let aiDoneFolderId     = null;
let aiTestingRunning   = false;

async function getAiTestingFolderId() {
  if (aiTestingFolderId) return aiTestingFolderId;
  const res  = await graphFetch(`/users/${BRANDON_EMAIL}/mailFolders?$select=id,displayName&$top=50`);
  const data = await res.json();
  const folder = (data.value || []).find(f => f.displayName === "AI testing");
  aiTestingFolderId = folder?.id || null;
  if (!aiTestingFolderId) console.warn("AI testing: 'AI testing' folder not found in Brandon's mailbox");
  return aiTestingFolderId;
}

async function getAiDoneFolderId() {
  if (aiDoneFolderId) return aiDoneFolderId;
  const res  = await graphFetch(`/users/${BRANDON_EMAIL}/mailFolders?$select=id,displayName&$top=50`);
  const data = await res.json();
  const folder = (data.value || []).find(f => f.displayName === "AI done");
  aiDoneFolderId = folder?.id || null;
  if (!aiDoneFolderId) console.warn("AI testing: 'AI done' folder not found in Brandon's mailbox");
  return aiDoneFolderId;
}

async function processAiTestingEmails() {
  if (aiTestingRunning) return;
  aiTestingRunning = true;
  try {
    console.log("AI testing: checking...");
    const folderId = await getAiTestingFolderId();
    if (!folderId) {
      console.log("AI testing: folder not found, skipping");
      return;
    }

    const res  = await graphFetch(
      `/users/${BRANDON_EMAIL}/mailFolders/${folderId}/messages` +
      `?$select=id,subject,body,categories,internetMessageId` +
      `&$expand=attachments($select=id,name,contentType,size)` +
      `&$top=50`
    );
    const data = await res.json();
    if (!res.ok) throw new Error(`Graph error ${res.status}: ${JSON.stringify(data?.error)}`);

    const messages = (data.value || []).filter(m => !/^re:/i.test(m.subject));
    console.log(`AI testing: ${messages.length} email(s) to process`);

    for (const message of messages) {
      try {
        const { result } = await processMessage(message, BRANDON_EMAIL);
        console.log("AI testing result:", result);

        const confidence = result["confidence"] ?? null;
        const aiNotes    = result["notes"] || null;
        const skipKeys   = new Set(["confidence", "notes"]);
        const rows = Object.entries(result)
          .filter(([k]) => !skipKeys.has(k))
          .map(([k, v]) => `<tr><td style="padding:3px 16px 3px 0;font-weight:bold;vertical-align:top">${k}</td><td style="padding:3px 0">${v ?? "<em>not found</em>"}</td></tr>`)
          .join("");

        const confidenceColour = confidence >= 0.8 ? "#2e7d32" : confidence >= 0.5 ? "#e65100" : "#c62828";
        const confidenceHtml = confidence !== null
          ? `<p style="font-family:sans-serif;font-size:14px"><strong>Confidence:</strong> <span style="color:${confidenceColour};font-weight:bold">${confidence}</span></p>`
          : "";
        const notesHtml = aiNotes
          ? `<p style="font-family:sans-serif;font-size:14px;color:#555"><strong>AI Notes:</strong> ${aiNotes}</p>`
          : "";

        const sendRes = await graphFetch(`/users/${BRANDON_EMAIL}/messages/${message.id}/reply`, {
          method: "POST",
          body: JSON.stringify({
            message: {
              toRecipients: [{ emailAddress: { address: BRANDON_EMAIL } }],
              body: {
                contentType: "HTML",
                content: `${confidenceHtml}${notesHtml}<p><strong>AI extracted the following from this work order:</strong></p><table style="font-family:sans-serif;font-size:14px;border-collapse:collapse">${rows}</table>`,
              },
            },
          }),
        });
        if (!sendRes.ok) {
          const err = await sendRes.json().catch(() => ({}));
          throw new Error(`Send failed ${sendRes.status}: ${JSON.stringify(err?.error)}`);
        }

        // Move to AI done folder so it won't be reprocessed
        const doneFolderId = await getAiDoneFolderId();
        if (doneFolderId) {
          await graphFetch(`/users/${BRANDON_EMAIL}/messages/${message.id}/move`, {
            method: "POST",
            body: JSON.stringify({ destinationId: doneFolderId }),
          });
        }

        console.log("AI testing: replied to", message.subject);
      } catch (err) {
        console.error("AI testing error:", message.subject, err.message);
      }
    }
  } catch (err) {
    console.error("AI testing poll error:", err.message);
  } finally {
    aiTestingRunning = false;
  }
}

// POLL LOOP
// ================================================================
async function setJobStatus(mailbox, messageId, currentCategories, newStatus) {
  const updated = [
    ...currentCategories.filter(c => !STATUS_CATEGORIES.includes(c)),
    newStatus,
  ];
  await graphFetch(`/users/${mailbox}/messages/${messageId}`, {
    method: "PATCH",
    body: JSON.stringify({ categories: updated }),
  });
  return updated;
}

async function pollInbox(mailbox) {
  const filter = encodeURIComponent(
    `categories/any(c:c eq '${TRIGGER_CATEGORY}')` +
    ` and not categories/any(c:c eq '${CLIENT_NOT_FOUND_CATEGORY}')` +
    ` and not categories/any(c:c eq '${PROCESSING_CATEGORY}')` +
    ` and not categories/any(c:c eq '${READING_EMAIL_CATEGORY}')` +
    ` and not categories/any(c:c eq '${SENDING_TO_AI_CATEGORY}')` +
    ` and not categories/any(c:c eq '${CREATING_JOB_CATEGORY}')`
  );

  const res  = await graphFetch(
    `/users/${mailbox}/mailFolders/inbox/messages` +
    `?$filter=${filter}` +
    `&$select=id,subject,body,categories` +
    `&$expand=attachments($select=id,name,contentType,size)` +
    `&$top=10`
  );
  const data = await res.json();
  if (!res.ok) throw new Error(`Graph API error ${res.status}: ${JSON.stringify(data?.error || data)}`);

  const messages = (data.value || []).filter(m =>
    !m.categories.some(c => c.startsWith("Job created"))
  );
  console.log(`Poll (${mailbox}): ${messages.length} email(s) found`);

  // Temp debug — show categories on recent inbox emails
  const dbg = await (await graphFetch(`/users/${mailbox}/mailFolders/inbox/messages?$select=subject,categories&$top=10&$orderby=receivedDateTime desc`)).json();
  for (const m of dbg.value || []) console.log("  >>", JSON.stringify(m.categories), m.subject?.slice(0, 60));

  for (const message of messages) {
    let currentCategories = await setJobStatus(mailbox, message.id, message.categories, READING_EMAIL_CATEGORY);
    console.log("Reading:", message.subject);

    try {
      const onStatus = async (status) => {
        currentCategories = await setJobStatus(mailbox, message.id, currentCategories, status);
      };

      const { result, rawEmail } = await processMessage(message, mailbox, onStatus);
      console.log("AI RESULT:", result);

      await onStatus(CREATING_JOB_CATEGORY);
      const { jobNumber } = await createArofloJob(result, rawEmail);

      // Success — strip all status tags, add job tag
      const jobTag = `Job created - ${jobNumber}`;
      currentCategories = await setJobStatus(mailbox, message.id, currentCategories, jobTag);
      console.log("Tagged as done:", message.subject);
    } catch (err) {
      console.error("Error processing message:", message.subject, err.message);
      if (err.message.startsWith("Client not found")) {
        await setJobStatus(mailbox, message.id, currentCategories, CLIENT_NOT_FOUND_CATEGORY);
        console.log("Tagged as client not found:", message.subject);
      } else {
        // Leave "Processing" on — remove it in Outlook to retry
        await setJobStatus(mailbox, message.id, currentCategories, PROCESSING_CATEGORY);
      }
    }
  }
}

async function pollEmails() {
  if (pollRunning) {
    console.log("Poll skipped — previous run still in progress");
    return;
  }
  pollRunning = true;
  try {
    await pollInbox(process.env.GRAPH_RECIPIENT);
    await pollInbox(BRANDON_EMAIL);
  } catch (err) {
    console.error("Poll error:", err);
  } finally {
    pollRunning = false;
  }
}

// ================================================================
// TEMP: List Aroflo task types — GET /task-types
// ================================================================
app.get("/task-types", async (req, res) => {
  try {
    const data = await arofloGet("zone=tasktypes&page=1");
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ================================================================
// TEMP: Test note posting — GET /test-note?job=103245
// ================================================================
app.get("/test-note", async (req, res) => {
  const jobNumber = req.query.job;
  if (!jobNumber) return res.status(400).json({ error: "Pass ?job=..." });

  try {
    // 1. Find the taskId for the job number
    const taskZone = await arofloGet(
      "zone=tasks&where=" + encodeURIComponent(`and|jobnumber|=|${jobNumber}`) + "&page=1"
    );
    const taskRaw  = taskZone.tasks;
    const taskArr  = taskRaw ? (Array.isArray(taskRaw) ? taskRaw : [taskRaw]) : [];
    if (taskArr.length === 0) return res.json({ error: `Job ${jobNumber} not found in Aroflo` });
    const taskId   = taskArr[0].taskid;

    // 2. Fetch the latest email from inbox (read-only, no locking)
    const emailRes  = await graphFetch(
      `/users/${process.env.GRAPH_RECIPIENT}/mailFolders/inbox/messages` +
      `?$select=id,subject,body&$orderby=receivedDateTime desc&$top=1`
    );
    const emailData = await emailRes.json();
    const message   = emailData.value?.[0];
    if (!message) return res.json({ error: "No emails found in inbox" });

    const rawEmail  = message.body?.content || "";

    // Extract a sample href to debug SafeLinks decoding
    const sampleHref = (rawEmail.match(/href="([^"]{30,})"/i) || [])[1] || "none found";

    const noteHtml  = await emailHtmlForNote(rawEmail);

    // 3. Try posting note via zone=tasknotes
    const results = {};
    for (const zone of ["tasknotes", "notes"]) {
      try {
        const xml = zone === "tasknotes"
          ? `<tasknotes><tasknote><taskid>${taskId}</taskid><content><![CDATA[${noteHtml}]]></content></tasknote></tasknotes>`
          : `<notes><note><taskid>${taskId}</taskid><content><![CDATA[${noteHtml}]]></content></note></notes>`;
        const r = await arofloPost(`zone=${zone}&postxml=` + encodeURIComponent(xml));
        results[zone] = r ?? "returned undefined";
      } catch (err) {
        results[zone] = { error: err.message };
      }
    }

    // Also try adding note via task update
    try {
      const updateXml = `<tasks><task><taskid>${taskId}</taskid><notes><note><content><![CDATA[${noteHtml}]]></content></note></notes></task></tasks>`;
      const r = await arofloPost("zone=tasks&postxml=" + encodeURIComponent(updateXml));
      results["task_update"] = r ?? "returned undefined";
    } catch (err) {
      results["task_update"] = { error: err.message };
    }

    res.json({ taskId, jobNumber, emailSubject: message.subject, sampleHref, results });
  } catch (err) {
    res.json({ error: err.message });
  }
});

// ================================================================
// Manual trigger for AI testing check
// ================================================================
app.get("/run-ai-test", async (req, res) => {
  res.json({ status: "triggered" });
  await processAiTestingEmails();
});

// ================================================================
// Manual trigger to process Brandon's inbox immediately
// ================================================================
app.get("/run-from-inbox", async (req, res) => {
  if (pollRunning) {
    res.json({ status: "skipped", reason: "poll already running" });
    return;
  }
  res.json({ status: "triggered" });
  pollRunning = true;
  try {
    await pollInbox(BRANDON_EMAIL);
  } catch (err) {
    console.error("run-from-inbox error:", err.message);
  } finally {
    pollRunning = false;
  }
});

// TEMP: Test Rica forwarding — GET /test-rica
// ================================================================
app.get("/test-rica", async (req, res) => {
  try {
    // 1. Check what Rica emails exist in workorders inbox
    const filter = encodeURIComponent(`categories/any(c:c eq '${RICA_CATEGORY}')`);
    const listRes  = await graphFetch(
      `/users/${WORKORDERS_EMAIL}/mailFolders/inbox/messages` +
      `?$filter=${filter}&$select=id,subject,body,categories&$top=50`
    );
    const listData = await listRes.json();
    if (!listRes.ok) return res.json({ step: "list Rica", error: listData?.error });
    const emails = listData.value || [];

    // Also grab recent emails from workorders to show actual category names for debugging
    const recentRes  = await graphFetch(
      `/users/${WORKORDERS_EMAIL}/mailFolders/inbox/messages` +
      `?$select=subject,categories&$orderby=receivedDateTime desc&$top=20`
    );
    const recentData = await recentRes.json();
    const recentCats = (recentData.value || []).map(m => ({ subject: m.subject?.slice(0, 60), categories: m.categories }));

    if (emails.length === 0) return res.json({ found: 0, message: "No Rica-tagged emails found in workorders inbox", recentEmails: recentCats });

    // 2. Forward the first one using the /forward endpoint (proven to work)
    const email   = emails[0];
    const sendRes = await graphFetch(
      `/users/${WORKORDERS_EMAIL}/messages/${email.id}/forward`,
      {
        method: "POST",
        body: JSON.stringify({
          toRecipients: [{ emailAddress: { address: BRANDON_EMAIL } }],
        }),
      }
    );
    const fwdStatus = sendRes.status;
    const fwdBody   = sendRes.status !== 202 ? await sendRes.json().catch(() => null) : null;

    // 3. Wait then check Brandon's inbox for recent emails
    await new Promise(r => setTimeout(r, 6000));
    const since     = new Date(Date.now() - 120000).toISOString();
    const findRes   = await graphFetch(
      `/users/${BRANDON_EMAIL}/mailFolders/inbox/messages` +
      `?$filter=${encodeURIComponent(`receivedDateTime ge ${since}`)}` +
      `&$select=id,subject,categories&$orderby=receivedDateTime desc&$top=10`
    );
    const findData  = await findRes.json();

    res.json({
      ricaEmailsFound:  emails.map(e => ({ id: e.id, subject: e.subject })),
      forwardStatus:    fwdStatus,
      forwardError:     fwdBody,
      brandonRecentEmails: (findData.value || []).map(m => ({ subject: m.subject, categories: m.categories })),
      brandonInboxError:   findRes.ok ? null : findData?.error,
    });
  } catch (err) {
    res.json({ error: err.message });
  }
});

// ================================================================
// START
// ================================================================
app.listen(process.env.PORT || 3000, () => {
  console.log(`Server running — deployed ${new Date().toISOString()}`);
  pollEmails();
  forwardRicaEmails();
  processAiTestingEmails();
  setInterval(pollEmails, POLL_INTERVAL_MS);
  setInterval(forwardRicaEmails, POLL_INTERVAL_MS);
  setInterval(processAiTestingEmails, POLL_INTERVAL_MS);
});

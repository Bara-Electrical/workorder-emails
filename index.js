import express from "express";
import OpenAI from "openai";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { createHmac } from "crypto";

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
const DONE_CATEGORY             = "Job created";
const CLIENT_NOT_FOUND_CATEGORY = "Client not found";
const RICA_CATEGORY             = "Rica";
const WORKORDERS_EMAIL          = "workorders@baraelectrical.com.au";
const BRANDON_EMAIL             = "brandon.roberts@baraelectrical.com.au";
const POLL_INTERVAL_MS          = 5 * 60 * 1000;

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
  "Real Estate Aircon Maintenance":  "JCZaSyBSQCAgCg==", // Real-Estate Air Con Maintenance
  "Real Estate General Maintenance": "JCYqWyVQICAgCg==", // Real-Estate General Maintenance
};

const SUBSTATUS_MAP = {
  "EC1":                            "IyQ6SyYK", // Ready to schedule PPM1
  "AC1":                            "IyQ6SyYK", // Ready to schedule PPM1
  "AC2":                            "Iyc6LyUK", // Ready to schedule (Specialised)
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
    result["property-manager"] ? `Property Manager: ${result["property-manager"]}` : null,
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
    ${taskTypeId ? `<tasktype><tasktypeid>${taskTypeId}</tasktypeid></tasktype>`                  : ""}
    <client><clientid>${client.clientid}</clientid></client>
    ${pmContact ? `<contact><userid>${pmContact.userid}</userid></contact>` : ""}
    ${location ? `<location><locationid>${location.locationid}</locationid></location>` : ""}
    ${result.address && !location  ? `<sitename>${result.address}</sitename>`          : ""}
    <taskname>${taskName}</taskname>
    <description>${result["task-description"] || result["task-type"] || ""}</description>
    <duedate>${dueDate}</duedate>
    ${result["order-number"] ? `<custon>${result["order-number"]}</custon>` : ""}
    ${result["account-to"] ? `<customfields><customfield><name><![CDATA[ Account To: ]]></name><type><![CDATA[ text ]]></type><value><![CDATA[${result["account-to"]}]]></value></customfield></customfields>` : ""}
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

  const inserted  = pr?.inserts?.tasks;
  const task      = Array.isArray(inserted) ? inserted[0] : inserted;
  const taskId    = task?.taskid;
  const jobNumber = task?.jobnumber || taskId || "(see Aroflo)";
  console.log("AROFLO JOB CREATED — job number:", jobNumber);

  // Post notes separately — inline notes on task creation are not supported
  if (taskId && (notes || rawEmail)) {
    const noteItems = [
      notes    ? `<tasknote><taskid>${taskId}</taskid><content><![CDATA[${notes}]]></content></tasknote>` : "",
      rawEmail ? `<tasknote><taskid>${taskId}</taskid><content><![CDATA[${await emailHtmlForNote(rawEmail)}]]></content></tasknote>` : "",
    ].join("");
    const notesXml = `<tasknotes>${noteItems}</tasknotes>`;
    try {
      const notesZone = await arofloPost("zone=tasknotes&postxml=" + encodeURIComponent(notesXml));
      console.log("Notes posted:", notesZone.postresults?.inserttotal ?? "unknown");
    } catch (err) {
      console.warn("Notes post failed:", err.message);
    }
  }

  // Aroflo doesn't apply substatus on create — do a follow-up update
  if (taskId && substatusId) {
    const updateXml =
`<tasks>
  <task>
    <taskid>${taskId}</taskid>
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

  return zone;
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
  const cleaned = html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<img[^>]*>/gi, "")
    .trim();
  return await decodeWrappedLinks(cleaned);
}

// Search raw HTML href attributes before cleaning strips them.
// Handles SafeLinks → Inky → TAPI redirect chains.
function findWorkOrderLink(rawHtml) {
  const unescaped = rawHtml.replace(/&amp;/g, "&");
  const hrefs     = [...unescaped.matchAll(/href="([^"]+)"/gi)].map(m => m[1]);

  for (const href of hrefs) {
    if (!href.startsWith("https://")) continue;
    if (/tapihq\.com/i.test(href)) return href;
    if (/shared\.outlook\.inky\.com/i.test(href)) return href;
    if (/safelinks\.protection\.outlook\.com/i.test(href)) {
      try {
        if (/tapihq\.com|inky\.com/i.test(decodeURIComponent(href))) return href;
      } catch { /* malformed URL */ }
    }
  }

  // Fallback: plain-text TAPI URL
  const plainText = unescaped.replace(/<[^>]*>/g, " ");
  const m = plainText.match(/https:\/\/url\d+\.tapihq\.com\/ls\/click\S+/i);
  if (m) return m[0].split(/[>")\s]/)[0].trim();

  return null;
}

// ================================================================
// EMAIL PROCESSING
// ================================================================
async function processMessage(message) {
  console.log("WORK ORDER EMAIL RECEIVED:", message.subject);

  const attachments = message.attachments || [];
  const rawBody     = message.body?.content || "";
  let   textForAI   = message.body?.contentType === "html" ? cleanHtml(rawBody) : rawBody;

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
      const html = await response.text();
      console.log("RAW HTML LENGTH:", html.length);
      console.log("HTML PREVIEW:", html.slice(0, 300));
      textForAI = cleanHtml(html).slice(0, 50000);
      console.log("CLEAN TEXT LENGTH:", textForAI.length);
    } catch (err) {
      console.error("LINK FETCH ERROR:", err);
    }
  }

  // Work order PDF attachment
  const workorderAttachment = attachments.find(a => {
    const name = (a.name || "").toLowerCase();
    return name.includes("workorder") && name.endsWith(".pdf");
  });

  if (workorderAttachment) {
    console.log("FOUND WORK ORDER PDF:", workorderAttachment.name);
    const attachRes  = await graphFetch(
      `/users/${process.env.GRAPH_RECIPIENT}/messages/${message.id}/attachments/${workorderAttachment.id}`
    );
    const attachData = await attachRes.json();
    const data       = Uint8Array.from(atob(attachData.contentBytes), c => c.charCodeAt(0));
    const pdfText    = await extractPDF(data);
    textForAI = pdfText.replace(/\s+/g, " ").trim();
    console.log("USING PDF CONTENT FOR AI");
  } else if (!workOrderLink) {
    console.log("NO PDF OR LINK - USING EMAIL BODY");
    textForAI = textForAI.replace(/\s+/g, " ").trim();
  }

  console.log("TEXT LENGTH:", textForAI.length);
  console.log("ABOUT TO CALL AI");

  const responseAI = await openai.responses.create({
    model: "gpt-5-mini",
    text: { format: { type: "json_object" } },
    input: `
You are a work order extraction system for an electrical company.

CRITICAL RULES:
- tenant-name must ONLY come from Tenant Details section.
- property-manager must ONLY come from Property Manager section
- account-to must include ALL owners exactly as written, and always be owners c/o real estate.
- do NOT guess missing fields
- if missing return null
- task-description must be concise electrician job summary
- order-number is the job/work order number
- if you cant find the real estate name, it'll be in the account to after the owners name or after the c/o.
- make sure to get all tenants names and numbers, there is often more than one and seperate them with commas not an array.
- Use tenants mobile numbers over home numbers if there is both.
- Australian phone numbers always start with 0 (e.g. 0412 345 678). Always include the leading 0.
- Check the page title to determine the task type.
- if there is no tenant details look for other access eg. lockbox with location. Put this in the tenant-name field and leave the contact blank.


TASK TYPES:
EC1 = Electrical Compliance Check
AC1 = Aircon Servicing
AC2 = Deluxe Aircon Clean
Real Estate Aircon Maintenance = aircon related jobs
Real Estate General Maintenance = everything else

Return ONLY JSON:

{
  "task-type": "",
  "tenant-name": "",
  "tenant-contact": "",
  "address": "",
  "task-description": "",
  "real-estate": "",
  "property-manager": "",
  "account-to": "",
  "order-number": ""
}

TEXT:
${textForAI}
    `,
  });

  const parsed = JSON.parse(responseAI.output_text);

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

    for (const email of emails) {
      const fwdSubject = `AI testing - FW: ${email.subject}`;

      // Dedup: check the "AI testing" folder (where Outlook rule moves them),
      // falling back to inbox if the folder isn't found.
      const safeSubj    = fwdSubject.replace(/'/g, "''");
      const dedupFilter = encodeURIComponent(`subject eq '${safeSubj}'`);
      const dedupFolder = aiTestingFolder
        ? `/users/${BRANDON_EMAIL}/mailFolders/${aiTestingFolder}/messages`
        : `/users/${BRANDON_EMAIL}/mailFolders/inbox/messages`;
      const dedupRes    = await graphFetch(`${dedupFolder}?$filter=${dedupFilter}&$select=id&$top=1`);
      const dedupData   = await dedupRes.json();
      if (dedupData.value?.length > 0) continue;

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

      // Send from workorders to Brandon with saveToSentItems: false — leaves zero
      // trace in the workorders mailbox (not even Sent Items)
      const message = {
        subject:      fwdSubject,
        body:         email.body,
        toRecipients: [{ emailAddress: { address: BRANDON_EMAIL } }],
      };
      if (attachments.length) message.attachments = attachments;

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
let aiTestingFolderId = null;

async function getAiTestingFolderId() {
  if (aiTestingFolderId) return aiTestingFolderId;
  const res  = await graphFetch(`/users/${BRANDON_EMAIL}/mailFolders?$select=id,displayName&$top=50`);
  const data = await res.json();
  const folder = (data.value || []).find(f => f.displayName === "AI testing");
  aiTestingFolderId = folder?.id || null;
  if (!aiTestingFolderId) console.warn("AI testing: 'AI testing' folder not found in Brandon's mailbox");
  return aiTestingFolderId;
}

async function processAiTestingEmails() {
  try {
    const folderId = await getAiTestingFolderId();
    if (!folderId) return;

    const res  = await graphFetch(
      `/users/${BRANDON_EMAIL}/mailFolders/${folderId}/messages` +
      `?$filter=${encodeURIComponent("not categories/any(c:c eq 'AI Replied')")}` +
      `&$select=id,subject,body,categories` +
      `&$expand=attachments($select=id,name,contentType,size)` +
      `&$top=10`
    );
    const data = await res.json();
    if (!res.ok) throw new Error(`Graph error ${res.status}: ${JSON.stringify(data?.error)}`);

    const messages = (data.value || []).filter(m => !m.subject?.startsWith("RE:"));
    if (messages.length) console.log(`AI testing: ${messages.length} email(s) to process`);

    for (const message of messages) {
      try {
        const { result } = await processMessage(message);
        console.log("AI testing result:", result);

        const rows = Object.entries(result)
          .map(([k, v]) => `<tr><td style="padding:3px 16px 3px 0;font-weight:bold;vertical-align:top">${k}</td><td style="padding:3px 0">${v ?? "<em>not found</em>"}</td></tr>`)
          .join("");

        await graphFetch(`/users/${BRANDON_EMAIL}/messages/${message.id}/reply`, {
          method: "POST",
          body: JSON.stringify({
            message: {
              toRecipients: [{ emailAddress: { address: BRANDON_EMAIL } }],
              body: {
                contentType: "HTML",
                content: `<p><strong>AI extracted the following from this work order:</strong></p><table style="font-family:sans-serif;font-size:14px;border-collapse:collapse">${rows}</table>`,
              },
            },
          }),
        });

        // Mark as processed so it won't be picked up again
        await graphFetch(`/users/${BRANDON_EMAIL}/messages/${message.id}`, {
          method: "PATCH",
          body: JSON.stringify({ categories: [...(message.categories || []), "AI Replied"] }),
        });
        console.log("AI testing: replied to", message.subject);
      } catch (err) {
        console.error("AI testing error:", message.subject, err.message);
      }
    }
  } catch (err) {
    console.error("AI testing poll error:", err.message);
  }
}

// POLL LOOP
// ================================================================
async function pollEmails() {
  if (pollRunning) {
    console.log("Poll skipped — previous run still in progress");
    return;
  }
  pollRunning = true;
  try {
    const filter = encodeURIComponent(
      `categories/any(c:c eq '${TRIGGER_CATEGORY}') and not categories/any(c:c eq '${DONE_CATEGORY}') and not categories/any(c:c eq '${CLIENT_NOT_FOUND_CATEGORY}') and not categories/any(c:c eq '${PROCESSING_CATEGORY}')`
    );

    const res  = await graphFetch(
      `/users/${process.env.GRAPH_RECIPIENT}/mailFolders/inbox/messages` +
      `?$filter=${filter}` +
      `&$select=id,subject,body,categories` +
      `&$expand=attachments($select=id,name,contentType,size)` +
      `&$top=10`
    );
    const data = await res.json();
    if (!res.ok) throw new Error(`Graph API error ${res.status}: ${JSON.stringify(data?.error || data)}`);
    const messages = data.value || [];
    console.log(`Poll: ${messages.length} email(s) found`);

    // Temp debug — show categories on recent inbox emails
    const dbg = await (await graphFetch(`/users/${process.env.GRAPH_RECIPIENT}/mailFolders/inbox/messages?$select=subject,categories&$top=10&$orderby=receivedDateTime desc`)).json();
    for (const m of dbg.value || []) console.log("  >>", JSON.stringify(m.categories), m.subject?.slice(0, 60));

    for (const message of messages) {
      // Lock the email — keep "Bara AI", add "Processing" so it won't re-trigger.
      // If something fails, just remove "Processing" in Outlook to retry.
      const lockedCategories = [...message.categories, PROCESSING_CATEGORY];
      await graphFetch(`/users/${process.env.GRAPH_RECIPIENT}/messages/${message.id}`, {
        method: "PATCH",
        body: JSON.stringify({ categories: lockedCategories }),
      });
      console.log("Locked for processing:", message.subject);

      try {
        const { result, rawEmail } = await processMessage(message);
        console.log("AI RESULT:", result);

        await createArofloJob(result, rawEmail);

        // Success — remove "Processing", add "Job created", keep everything else
        const doneCategories = [
          ...lockedCategories.filter(c => c !== PROCESSING_CATEGORY),
          DONE_CATEGORY,
        ];
        await graphFetch(`/users/${process.env.GRAPH_RECIPIENT}/messages/${message.id}`, {
          method: "PATCH",
          body: JSON.stringify({ categories: doneCategories }),
        });
        console.log("Tagged as done:", message.subject);
      } catch (err) {
        console.error("Error processing message:", message.subject, err.message);
        if (err.message.startsWith("Client not found")) {
          const clientNotFoundCategories = [
            ...lockedCategories.filter(c => c !== PROCESSING_CATEGORY),
            CLIENT_NOT_FOUND_CATEGORY,
          ];
          await graphFetch(`/users/${process.env.GRAPH_RECIPIENT}/messages/${message.id}`, {
            method: "PATCH",
            body: JSON.stringify({ categories: clientNotFoundCategories }),
          });
          console.log("Tagged as client not found:", message.subject);
        }
        // Any other error: leave "Processing" on — remove it in Outlook to retry
      }
    }
  } catch (err) {
    console.error("Poll error:", err);
  } finally {
    pollRunning = false;
  }
}

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
  console.log("Server running");
  pollEmails();
  forwardRicaEmails();
  processAiTestingEmails();
  setInterval(pollEmails, POLL_INTERVAL_MS);
  setInterval(forwardRicaEmails, POLL_INTERVAL_MS);
  setInterval(processAiTestingEmails, POLL_INTERVAL_MS);
});

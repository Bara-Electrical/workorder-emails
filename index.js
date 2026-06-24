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

const TRIGGER_CATEGORY       = "Bara AI";
const DONE_CATEGORY          = "Job created";
const CLIENT_NOT_FOUND_CATEGORY = "Client not found";
const POLL_INTERVAL_MS       = 60 * 1000;

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
      Accept:              AROFLO_ACCEPT,
      Authorization:       auth,
      "af-hmac-signature": arofloSign("GET", params, ts),
      "af-iso-timestamp":  ts,
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
      Accept:              AROFLO_ACCEPT,
      Authorization:       auth,
      "af-hmac-signature": arofloSign("POST", body, ts),
      "af-iso-timestamp":  ts,
      "Content-Type":      "application/x-www-form-urlencoded",
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

// Fetch all clients (85 total, fits in one page) and find by name client-side.
// Tries progressively shorter variants of the real estate name.
async function findClient(realEstateName) {
  if (!realEstateName) return null;

  const zone    = await arofloGet("zone=clients&page=1");
  const clients = zone.clients || [];
  const arr     = Array.isArray(clients) ? clients : [clients];

  const nameLower  = realEstateName.toLowerCase();
  const beforePipe = realEstateName.split(/[|,]/)[0].trim().toLowerCase(); // "Ray White Cottesloe"
  const firstWord  = realEstateName.split(" ")[0].toLowerCase();           // "Ray"

  return (
    arr.find(c => c.clientname?.toLowerCase() === nameLower) ||
    arr.find(c => c.clientname?.toLowerCase().includes(beforePipe)) ||
    arr.find(c => c.clientname?.toLowerCase().includes(firstWord)) ||
    null
  );
}

// Find a location by street address, then update SiteContact/SitePhone if stale.
async function findOrUpdateLocation(address, tenantName, tenantContact) {
  if (!address) return null;

  // Strip unit: "1412/380 Murray Street, Perth WA" → "380 Murray Street"
  const streetPart = address.replace(/^\d+\//, "").split(",")[0].trim();

  let zone;
  try {
    zone = await arofloGet(
      "zone=locations" +
      "&where=" + encodeURIComponent(`and|locationname|like|${streetPart}`) +
      "&page=1"
    );
  } catch (err) {
    console.warn("Location search failed:", err.message);
    return null;
  }

  const raw = zone.locations;
  if (!raw) { console.log("No location found for:", streetPart); return null; }

  const location = Array.isArray(raw) ? raw[0] : raw;
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

async function createArofloJob(result) {
  console.log("CREATING AROFLO JOB...");

  const taskTypeId = TASK_TYPE_MAP[result["task-type"]];
  if (!taskTypeId) console.warn("Unknown task type:", result["task-type"]);

  const client = await findClient(result["real-estate"]);
  if (!client) throw new Error(`Client not found in Aroflo: "${result["real-estate"]}"`);
  console.log("CLIENT:", client.clientid, client.clientname);

  const location = await findOrUpdateLocation(
    result.address,
    result["tenant-name"],
    result["tenant-contact"]
  );

  const notes = [
    result["order-number"]     ? `Work Order: ${result["order-number"]}`          : null,
    result["tenant-name"]      ? `Tenant: ${result["tenant-name"]}`                : null,
    result["tenant-contact"]   ? `Tenant Contact: ${result["tenant-contact"]}`     : null,
    result["property-manager"] ? `Property Manager: ${result["property-manager"]}` : null,
    result["account-to"]       ? `Account To: ${result["account-to"]}`             : null,
  ].filter(Boolean).join("\n");

  const xml =
`<tasks>
  <task>
    ${taskTypeId                   ? `<tasktypeid>${taskTypeId}</tasktypeid>`          : ""}
    <clientid>${client.clientid}</clientid>
    ${location                     ? `<locationid>${location.locationid}</locationid>` : ""}
    ${result.address && !location  ? `<sitename>${result.address}</sitename>`          : ""}
    <description>${result["task-description"] || result["task-type"] || ""}</description>
    ${notes ? `<notes><note><content><![CDATA[${notes}]]></content></note></notes>` : ""}
  </task>
</tasks>`;

  const zone = await arofloPost("zone=tasks&postxml=" + encodeURIComponent(xml));
  console.log("AROFLO JOB CREATED:", JSON.stringify(zone));
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

  return JSON.parse(responseAI.output_text);
}

// ================================================================
// POLL LOOP
// ================================================================
async function pollEmails() {
  try {
    const filter = encodeURIComponent(
      `categories/any(c:c eq '${TRIGGER_CATEGORY}') and not categories/any(c:c eq '${DONE_CATEGORY}') and not categories/any(c:c eq '${CLIENT_NOT_FOUND_CATEGORY}')`
    );

    const res  = await graphFetch(
      `/users/${process.env.GRAPH_RECIPIENT}/mailFolders/inbox/messages` +
      `?$filter=${filter}` +
      `&$select=id,subject,body,categories` +
      `&$expand=attachments($select=id,name,contentType,size)` +
      `&$top=10`
    );
    const data     = await res.json();
    const messages = data.value || [];

    if (messages.length) console.log(`Found ${messages.length} email(s) to process`);

    for (const message of messages) {
      try {
        const result = await processMessage(message);
        console.log("AI RESULT:", result);

        await createArofloJob(result);

        // Only tag as done after successful job creation
        await graphFetch(`/users/${process.env.GRAPH_RECIPIENT}/messages/${message.id}`, {
          method: "PATCH",
          body: JSON.stringify({ categories: [...message.categories, DONE_CATEGORY] }),
        });
        console.log("Tagged as done:", message.subject);
      } catch (err) {
        console.error("Error processing message:", message.subject, err.message);
        if (err.message.startsWith("Client not found")) {
          await graphFetch(`/users/${process.env.GRAPH_RECIPIENT}/messages/${message.id}`, {
            method: "PATCH",
            body: JSON.stringify({ categories: [...message.categories, CLIENT_NOT_FOUND_CATEGORY] }),
          });
          console.log("Tagged as client not found:", message.subject);
        }
      }
    }
  } catch (err) {
    console.error("Poll error:", err);
  }
}

// ================================================================
// START
// ================================================================
app.listen(process.env.PORT || 3000, () => {
  console.log("Server running");
  pollEmails();
  setInterval(pollEmails, POLL_INTERVAL_MS);
});

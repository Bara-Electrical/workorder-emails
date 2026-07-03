import "dotenv/config";
import express from "express";
import OpenAI from "openai";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { createHmac } from "crypto";
import { PACKAGE_TEMPLATES } from "./templates.js";
import Airtable from "airtable";

const REQUIRED_ENV = [
  "OPENAI_API_KEY",
  "GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET",
  "UENCODED", "PENCODED", "ORGENCODED", "SECRET_KEY",
];
for (const key of REQUIRED_ENV) {
  if (!process.env[key]) {
    console.error(`${key} is not set`);
    process.exit(1);
  }
}

let airtableBase = null;
if (process.env.AIRTABLE_API_KEY && process.env.AIRTABLE_BASE_ID) {
  airtableBase = new Airtable({ apiKey: process.env.AIRTABLE_API_KEY }).base(process.env.AIRTABLE_BASE_ID);
} else {
  console.warn("AIRTABLE_API_KEY or AIRTABLE_BASE_ID not set — activity and AI logging disabled");
}

// Known sender name → Aroflo client name mappings
const CLIENT_NAME_MAP = {
  "warilla pty ltd": "Peter Kuhne Real Estate",
  "warrilla pty ltd": "Peter Kuhne Real Estate",
  "professionals cannington": "Professionals BW Backhouse & Associates",
  "bw backhouse & associates": "Professionals BW Backhouse & Associates",
  "acton belle property south perth & victoria park": "Acton | Belle Property South Perth and Victoria Park",
  "leifield": "Leifield Real Estate",
  "leifield - wa": "Leifield Real Estate",
};

// Sender email domain → Aroflo client name (fallback when AI can't extract name from compound domains)
const EMAIL_DOMAIN_MAP = {
  "platinumelectricians.com.au": "Platinum Electricians",
};

const TRIGGER_CATEGORY          = "Bara AI";
const PROCESSING_CATEGORY       = "Processing";
const FAILED_CATEGORY           = "Failed";
const CLIENT_NOT_FOUND_CATEGORY = "Client not found";
const NO_ADDRESS_CATEGORY       = "No address";
const READING_EMAIL_CATEGORY    = "Reading Email";
const SENDING_TO_AI_CATEGORY    = "Sending to AI";
const CREATING_JOB_CATEGORY     = "Creating Job";
const RICA_CATEGORY             = "Rica";
const WORKORDERS_EMAIL          = "workorders@baraelectrical.com.au";
const BRANDON_EMAIL             = "brandon.roberts@baraelectrical.com.au";
const POLL_INTERVAL_MS          = 5 * 60 * 1000;

// All transient status categories — stripped when transitioning to next stage
const STATUS_CATEGORIES = [
  PROCESSING_CATEGORY, FAILED_CATEGORY, READING_EMAIL_CATEGORY,
  SENDING_TO_AI_CATEGORY, CREATING_JOB_CATEGORY,
  CLIENT_NOT_FOUND_CATEGORY, NO_ADDRESS_CATEGORY,
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

// In-memory client cache: lowercase clientname → client object.
// Populated at startup and updated via the /aroflo-webhook endpoint.
const clientCache = new Map();

async function loadClientCache() {
  // where=clientid!=0 forces Aroflo to return the full client list.
  // Without a WHERE clause the API silently limits to recently active clients only.
  const WHERE = encodeURIComponent("and|clientid|!=|0");
  let page = 1, loaded = 0;
  try {
    while (true) {
      const zone = await arofloGet(`zone=clients&where=${WHERE}&page=${page}`);
      const raw  = zone?.clients;
      if (!raw) break;
      const arr  = Array.isArray(raw) ? raw : [raw];
      for (const c of arr) clientCache.set(c.clientname.toLowerCase(), c);
      loaded += arr.length;
      const current = parseInt(zone.currentpageresults ?? 0);
      const max     = parseInt(zone.maxpageresults ?? 500);
      if (current < max) break;
      page++;
    }
    console.log(`Client cache loaded: ${clientCache.size} unique clients across ${page} page(s)`);
  } catch (err) {
    console.error("Failed to load client cache — fuzzy matching unavailable:", err.message);
  }
}

// Fetch a client's locations and contacts in a single call (join=locations,contacts)
// instead of two separate lookups.
async function findLocationsAndContacts(clientId) {
  try {
    const zone = await arofloGet(
      "zone=clients" +
      "&join=" + encodeURIComponent("locations,contacts") +
      "&where=" + encodeURIComponent(`and|clientid|=|${clientId}`) +
      "&where=" + encodeURIComponent("and|archived|=|false") +
      "&page=1"
    );
    const raw    = zone.clients;
    const client = Array.isArray(raw) ? raw[0] : raw;
    const locs     = client?.locations;
    const contacts = client?.contacts;
    return {
      locations: locs ? (Array.isArray(locs) ? locs : [locs]) : [],
      contacts:  contacts ? (Array.isArray(contacts) ? contacts : [contacts]) : [],
    };
  } catch (err) {
    console.warn("Location/contact search failed:", err.message);
    return { locations: [], contacts: [] };
  }
}

// Match a PM contact by name from an already-fetched contacts array.
function matchContact(contacts, pmName) {
  if (!pmName) return null;
  const nameLower = pmName.toLowerCase();
  return contacts.find(c => `${c.givennames} ${c.surname}`.toLowerCase().includes(nameLower)) || null;
}

// Search for a client by name.
// 1. Exact match against local cache (no API call needed).
// 2. Starts-with fuzzy match against local cache — only if exactly one result.
// 3. Falls back to Aroflo API exact match if cache is empty (not yet loaded).
async function findClient(realEstateName) {
  if (!realEstateName) return null;

  const baseName    = realEstateName.split(/[|,]/)[0].trim();
  const strippedName = baseName.replace(/\s+(?:Pty\.?\s*)?(?:Ltd\.?|Limited|Inc\.?|LLC)\.?$/i, "").trim();

  // Cache lookup — check exact match on each candidate
  if (clientCache.size > 0) {
    const candidates = [
      realEstateName,
      baseName,
      strippedName,
      baseName.split(" ")[0],
    ].filter((v, i, a) => v && a.indexOf(v) === i);

    for (const name of candidates) {
      const hit = clientCache.get(name.toLowerCase());
      if (hit) return hit;
    }

    // Bidirectional starts-with fuzzy match — handles "Driven Property Group Pty Ltd" → "Driven Property Group"
    const query   = baseName.toLowerCase();
    const matches = [...clientCache.values()].filter(c => {
      const name = c.clientname.toLowerCase();
      return name.startsWith(query) || query.startsWith(name);
    });
    if (matches.length === 1) {
      console.log(`Fuzzy client match: "${realEstateName}" → "${matches[0].clientname}"`);
      return matches[0];
    }
    if (matches.length > 1) {
      console.warn(`Ambiguous client name "${realEstateName}" — ${matches.length} cache matches: ${matches.map(c => c.clientname).join(", ")}`);
    }
    return null;
  }

  // Cache not loaded yet — fall back to API
  const candidates = [
    realEstateName,
    baseName,
    strippedName,
    baseName.split(" ")[0],
  ].filter((v, i, a) => v && a.indexOf(v) === i);

  for (const name of candidates) {
    const zone = await arofloGet(
      "zone=clients&where=" + encodeURIComponent(`and|clientname|=|${name}`) + "&page=1"
    );
    const raw = zone.clients;
    if (!raw) continue;
    const arr = Array.isArray(raw) ? raw : [raw];
    if (arr.length > 0) return arr[0];
  }

  return null;
}

function parseAustralianAddress(address) {
  const parts = address.split(",").map(p => p.trim());
  const street = parts[0] || "";
  const rest   = parts.slice(1).join(", ").trim();
  // Optional comma before state handles "Perth, WA 6000" and "Perth WA 6000".
  // Explicit state list handles title-cased abbreviations like "Wa" or "Nsw".
  const match  = rest.match(/^(.*?),?\s+(NSW|VIC|QLD|SA|WA|TAS|ACT|NT)(?:[,\s]+(\d{4}))?$/i);
  return match
    ? { street, suburb: match[1].trim(), state: match[2].toUpperCase(), postcode: match[3] || "" }
    : { street, suburb: rest, state: "", postcode: "" };
}

async function geocodeAddress(address) {
  const key = process.env.HERE_API_KEY;
  if (!key) { console.warn("HERE_API_KEY not set — skipping geocode"); return null; }
  try {
    const url  = "https://geocode.search.hereapi.com/v1/geocode?q=" +
                 encodeURIComponent(address) + "&in=countryCode:AUS&apiKey=" + key;
    const res  = await fetch(url);
    const data = await res.json();
    const item = data?.items?.[0];
    if (item?.position) return { lat: item.position.lat, lon: item.position.lng };
  } catch (err) {
    console.warn("Geocode failed:", err.message);
  }
  return null;
}

async function createLocation(clientId, address, tenantName, tenantContact, tenantEmail) {
  const { street, suburb, state, postcode } = parseAustralianAddress(address);
  const coords = await geocodeAddress(address);
  const xml =
`<clients><client>
  <clientid>${clientId}</clientid>
  <locations><location>
    <locationname><![CDATA[${street}]]></locationname>
    <suburb><![CDATA[${suburb}]]></suburb>
    <state><![CDATA[${state}]]></state>
    <postcode><![CDATA[${postcode}]]></postcode>
    <country><![CDATA[Australia]]></country>
    ${coords ? `<gpslat>${coords.lat}</gpslat><gpslong>${coords.lon}</gpslong>` : ""}
    ${tenantName    ? `<sitecontact><![CDATA[${tenantName}]]></sitecontact>`  : ""}
    ${tenantContact ? `<sitephone><![CDATA[${tenantContact}]]></sitephone>`   : ""}
    ${tenantEmail   ? `<siteemail><![CDATA[${tenantEmail}]]></siteemail>`     : ""}
  </location></locations>
</client></clients>`;
  const createZone = await arofloPost("zone=clients&postxml=" + encodeURIComponent(xml));
  console.log("Location create response:", JSON.stringify(createZone?.postresults));

  // Aroflo returns the new locationid under updates.clients[0].locations[0]
  const pr          = createZone?.postresults;
  const updClients  = pr?.updates?.clients;
  const updClient   = Array.isArray(updClients) ? updClients[0] : updClients;
  const updLocs     = updClient?.locations;
  const newId       = (Array.isArray(updLocs) ? updLocs[0] : updLocs)?.locationid;
  if (newId) {
    console.log("Location created:", newId, address);
    return { locationid: newId, locationname: address };
  }

  // Fall back: re-query to find the location we just created
  console.log("locationid not in response — fetching newly created location");
  const zone = await arofloGet(
    "zone=clients" +
    "&where=" + encodeURIComponent(`and|clientid|=|${clientId}`) +
    "&join=locations"
  );
  const client = Array.isArray(zone.clients) ? zone.clients[0] : zone.clients;
  const all = client?.locations ? (Array.isArray(client.locations) ? client.locations : [client.locations]) : [];
  const streetPart = address.replace(/^\d+\//, "").split(",")[0].trim().toLowerCase();
  const found = all.find(l => l.locationname?.toLowerCase().includes(streetPart));
  if (found) {
    console.log("Location found after creation:", found.locationid, found.locationname);
    return found;
  }

  console.warn("Location created but could not retrieve locationid");
  return null;
}

// Aroflo silently truncates SiteContact/SitePhone to 50 characters each (confirmed
// empirically — no validation error, it just clips). For multi-tenant jobs this can
// cut off a name or number mid-word. Keep the same number of tenants in both fields
// (so name[i] still lines up with phone[i]) and report anything dropped separately.
const SITE_FIELD_LIMIT = 50;
function fitTenantFields(tenantName, tenantContact, maxLen = SITE_FIELD_LIMIT) {
  const names  = tenantName    ? tenantName.split(",").map(s => s.trim())    : [];
  const phones = tenantContact ? tenantContact.split(",").map(s => s.trim()) : [];
  const total  = Math.max(names.length, phones.length);
  const fitsWithin = (items, count) => items.slice(0, count).join(", ").length <= maxLen;

  let keptCount = total;
  while (keptCount > 0 &&
         ((names.length  && !fitsWithin(names, keptCount)) ||
          (phones.length && !fitsWithin(phones, keptCount)))) {
    keptCount--;
  }

  // Single tenant longer than the limit on its own — hard-truncate rather than drop it entirely.
  if (keptCount === 0 && total > 0) {
    return {
      keptName:     names[0]  ? names[0].slice(0, maxLen)  : null,
      keptPhone:    phones[0] ? phones[0].slice(0, maxLen) : null,
      overflowName:  names.slice(1),
      overflowPhone: phones.slice(1),
      truncated: true,
    };
  }

  return {
    keptName:      names.slice(0, keptCount).join(", ")  || null,
    keptPhone:      phones.slice(0, keptCount).join(", ") || null,
    overflowName:   names.slice(keptCount),
    overflowPhone:  phones.slice(keptCount),
    truncated: keptCount < total,
  };
}

// Find a location by street address from an already-fetched locations array, then
// update SiteContact/SitePhone if stale, or create the location if it doesn't exist yet.
// Matches by street address locally — the locationname|like WHERE clause is not
// supported by Aroflo.
async function findOrUpdateLocation(clientId, locations, address, tenantName, tenantContact, tenantEmail) {
  if (!address) return null;

  // Strip unit prefix: "1412/380 Murray Street, Perth WA" → "380 Murray Street"
  const streetPart = address.replace(/^\d+\//, "").split(",")[0].trim().toLowerCase();
  const incomingUnit = address.match(/^(\d+)\//)?.[1] || null;

  const forClient = locations;
  const active = forClient.filter(l => l.archived?.toUpperCase() !== "TRUE");
  console.log(`Location search — client ${clientId}: ${forClient.length} location(s) (${active.length} active), searching for "${streetPart}"`);
  // A street-only match isn't enough when a building has multiple numbered units on
  // file — "10/27 X" contains "27 x" as a substring, so it would wrongly match an
  // incoming "9/27 X" and silently attach the job to a different unit's tenant.
  // If both the incoming address and the candidate location have a unit number,
  // they must match exactly.
  const location = active.find(l => {
    if (!l.locationname?.toLowerCase().includes(streetPart)) return false;
    const storedUnit = l.locationname?.match(/^(\d+)\//)?.[1] || null;
    if (incomingUnit && storedUnit && incomingUnit !== storedUnit) return false;
    return true;
  });

  if (!location) {
    console.log("No location matching:", streetPart, "— creating new location:", address);
    try {
      return await createLocation(clientId, address, tenantName, tenantContact, tenantEmail);
    } catch (err) {
      console.warn("Location creation failed:", err.message);
      return null;
    }
  }

  console.log("FOUND LOCATION:", location.locationid, location.locationname);

  if (tenantName || tenantContact || tenantEmail) {
    // Must be wrapped in <clients><client> — a bare <locations><location> POST to
    // zone=locations returns status "0" with no error but silently does not apply.
    const xml =
`<clients><client>
  <clientid>${clientId}</clientid>
  <locations><location>
    <locationid>${location.locationid}</locationid>
    ${tenantName    ? `<sitecontact><![CDATA[${tenantName}]]></sitecontact>`  : ""}
    ${tenantContact ? `<sitephone><![CDATA[${tenantContact}]]></sitephone>`   : ""}
    ${tenantEmail   ? `<siteemail><![CDATA[${tenantEmail}]]></siteemail>`     : ""}
  </location></locations>
</client></clients>`;
    try {
      const updateRes = await arofloPost("zone=clients&postxml=" + encodeURIComponent(xml));
      console.log("Tenant details updated:", JSON.stringify(updateRes?.postresults));
    } catch (err) {
      console.warn("Tenant update failed:", err.message);
    }
  }

  return location;
}

async function logActivity(action, jobNumber) {
  if (!airtableBase) return;
  try {
    await airtableBase("Activity Log").create([{
      fields: { "Action": action, "Job number": jobNumber || null, "Department": "Admin" },
    }]);
  } catch (err) {
    console.warn("Airtable activity log failed:", err.message);
  }
}

async function logAiOutput(result, emailSubject) {
  if (!airtableBase) return;
  try {
    await airtableBase("Work Order AI Log").create([{
      fields: {
        "Email Subject":     emailSubject || null,
        "Task Type":         result["task-type"] || null,
        "Package":           result["package"] || null,
        "Address":           result["address"] || null,
        "Real Estate":       result["real-estate"] || null,
        "Property Manager":  result["property-manager"] || null,
        "Tenant Name":       result["tenant-name"] || null,
        "Tenant Contact":    result["tenant-contact"] || null,
        "Order Number":      result["order-number"] || null,
        "Account To":        result["account-to"] || null,
        "Access Details":    result["access-details"] || null,
        "Expenditure Limit": result["expenditure-limit"] || null,
        "Task Description":  result["task-description"] || null,
        "Confidence":        result["confidence"] != null ? Number(result["confidence"]) : null,
        "AI Notes":          result["notes"] || null,
      },
    }]);
  } catch (err) {
    console.warn("Airtable AI log failed:", err.message);
  }
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

async function createArofloJob(result, rawEmail, pdfAttachment = null, emailMeta = null, imageAttachments = []) {
  console.log("CREATING AROFLO JOB...");
  const warnings = [];

  if (!result.address) throw new Error("No address found in work order");

  const taskTypeId  = TASK_TYPE_MAP[result["task-type"]];
  const substatusId = SUBSTATUS_MAP[result["task-type"]] || "Iyc6LyYK"; // default: Ready to schedule
  if (!taskTypeId) {
    const detail = `Unknown task type: "${result["task-type"]}" — task created without a task type`;
    console.warn(detail);
    warnings.push({ tag: "Unknown task type", detail });
  }

  const realEstate = CLIENT_NAME_MAP[result["real-estate"]?.toLowerCase()] || result["real-estate"];
  console.log(`Client lookup — AI extracted real-estate: "${result["real-estate"]}", resolved to: "${realEstate}", from: "${emailMeta?.from}"`);
  let client = await findClient(realEstate);
  let clientFoundVia = "name";
  if (!client && emailMeta?.from) {
    const domain = emailMeta.from.split("@")[1];
    const domainName = domain && EMAIL_DOMAIN_MAP[domain.toLowerCase()];
    console.log(`Client not found by name — domain: "${domain}", domain map hit: "${domainName || "none"}"`);
    if (domainName) {
      client = await findClient(domainName);
      if (!client) {
        // Client not in cache (may be a supplier/subcontractor type) — try live Aroflo API
        console.log(`Cache miss for "${domainName}" — trying live Aroflo API lookup`);
        const zone = await arofloGet(`zone=clients&where=${encodeURIComponent(`and|clientname|=|${domainName}`)}&page=1`);
        const raw = zone?.clients;
        if (raw) client = Array.isArray(raw) ? raw[0] : raw;
      }
      clientFoundVia = `email domain (${domainName})`;
    }
  }
  if (!client) throw new Error(`Client not found in Aroflo: name="${realEstate}", from="${emailMeta?.from}"`);
  console.log(`CLIENT (via ${clientFoundVia}):`, client.clientid, client.clientname);

  const { locations, contacts } = await findLocationsAndContacts(client.clientid);

  const tenantFit = fitTenantFields(result["tenant-name"], result["tenant-contact"]);
  let additionalTenantNote = null;
  if (tenantFit.truncated) {
    const rows = Math.max(tenantFit.overflowName.length, tenantFit.overflowPhone.length);
    const lines = [];
    for (let i = 0; i < rows; i++) {
      const line = [tenantFit.overflowName[i], tenantFit.overflowPhone[i]].filter(Boolean).join(", ");
      if (line) lines.push(`Additional tenant - ${line}`);
    }
    additionalTenantNote = lines.join("<br/>");
    const detail = `SiteContact/SitePhone are capped at ${SITE_FIELD_LIMIT} characters by Aroflo — full list: "${result["tenant-name"] || ""}" / "${result["tenant-contact"] || ""}"`;
    console.warn(detail);
    warnings.push({ tag: "Tenant details truncated", detail });
  }

  const location = await findOrUpdateLocation(
    client.clientid,
    locations,
    result.address,
    tenantFit.keptName,
    tenantFit.keptPhone,
    result["tenant-email"]
  );
  if (!location && result.address) {
    const detail = `Location not linked for "${result.address}" — address used as site name fallback`;
    console.warn(detail);
    warnings.push({ tag: "Location not linked", detail });
  }

  const pmContact = matchContact(contacts, result["property-manager"]);
  if (pmContact) {
    console.log("PM CONTACT:", pmContact.contactid, pmContact.contactname);
  } else if (result["property-manager"]) {
    const detail = `PM contact not found in Aroflo: "${result["property-manager"]}"`;
    console.warn(detail);
    warnings.push({ tag: "PM not in Aroflo", detail });
  }

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
    d.setDate(d.getDate() + 7);
    return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
  })();

  const { street: taskStreet, suburb: taskSuburb } = parseAustralianAddress(result.address || "");
  const taskName = [taskStreet, taskSuburb]
    .filter(Boolean)
    .join(" ")
    .replace(/,?\s*(?:NSW|VIC|QLD|SA|WA|TAS|ACT|NT)\b\s*\d{0,4}\s*$/i, "")
    .trim();

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
    ${(result["account-to"] || realEstate) ? `<customfields><customfield><name><![CDATA[ Account To: ]]></name><type><![CDATA[ text ]]></type><value><![CDATA[${result["account-to"] || realEstate}]]></value></customfield></customfields>` : ""}
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

  // Upload PDF and any photo attachments to SharePoint
  let oneDriveUrl = null;
  const photos = [];
  if (jobNumber !== "(see Aroflo)") {
    if (pdfAttachment) {
      try {
        oneDriveUrl = await uploadWorkOrderToOneDrive(jobNumber, pdfAttachment.name, pdfAttachment.data);
        console.log("PDF uploaded to OneDrive:", oneDriveUrl);
      } catch (err) {
        console.warn("OneDrive upload failed:", err.message);
        warnings.push({ tag: "PDF upload failed", detail: err.message });
      }
    }
    if (imageAttachments.length > 0) {
      const driveId = await getSharepointDriveId();
      const photoResults = await Promise.all(
        imageAttachments.map(async img => {
          try {
            const item = await uploadPhotoToOneDrive(jobNumber, img.name, img.data, img.contentType);
            const thumbnailUrl = await getThumbnailUrl(driveId, item.id).catch(() => null);
            console.log("Photo uploaded:", img.name);
            return { name: img.name, webUrl: item.webUrl, thumbnailUrl };
          } catch (err) {
            console.warn("Photo upload failed:", img.name, err.message);
            warnings.push({ tag: "Photo upload failed", detail: `${img.name}: ${err.message}` });
            return null;
          }
        })
      );
      photos.push(...photoResults.filter(Boolean));
    }
  }

  // Post the original email as a note and set the substatus in one combined task
  // update (Aroflo doesn't apply substatus on create, so a follow-up write is
  // always needed — piggyback the note content on the same call).
  if (confirmedTaskId) {
    let noteHtml = null;
    if (rawEmail) {
      try {
        noteHtml = await emailHtmlForNote(rawEmail, oneDriveUrl, emailMeta);
      } catch (err) {
        const detail = `Email note not posted to job: ${err.message}`;
        console.warn(detail);
        warnings.push({ tag: "Note not posted", detail });
      }
    } else {
      warnings.push({ tag: "Note not posted", detail: "No email content available — note not posted to job" });
    }

    const photoGalleryHtml = photos.length > 0 ? buildPhotoGalleryNote(photos) : null;

    const notesXml = [
      noteHtml             ? `<note><content><![CDATA[${noteHtml}]]></content></note>`             : "",
      photoGalleryHtml     ? `<note><content><![CDATA[${photoGalleryHtml}]]></content></note>`      : "",
      additionalTenantNote ? `<note><content><![CDATA[${additionalTenantNote}]]></content></note>` : "",
    ].join("");

    if (notesXml || substatusId) {
      const updateXml =
`<tasks>
  <task>
    <taskid>${confirmedTaskId}</taskid>
    ${substatusId ? `<status>not started</status><substatus><substatusid>${substatusId}</substatusid></substatus>` : ""}
    ${notesXml ? `<notes>${notesXml}</notes>` : ""}
  </task>
</tasks>`;
      try {
        const upZone = await arofloPost("zone=tasks&postxml=" + encodeURIComponent(updateXml));
        const upPr   = upZone.postresults;
        if (Number(upPr?.updatetotal ?? 0) > 0) {
          console.log("Task update applied — note:", !!noteHtml, "additional tenant note:", !!additionalTenantNote, "substatus:", substatusId || "n/a");
        } else {
          if (noteHtml) warnings.push({ tag: "Note not posted", detail: "Combined task update did not apply" });
          if (substatusId) warnings.push({ tag: "Substatus failed", detail: "Combined task update did not apply — job may need manual scheduling status update" });
        }
      } catch (err) {
        if (noteHtml) warnings.push({ tag: "Note not posted", detail: `Email note not posted to job: ${err.message}` });
        if (substatusId) warnings.push({ tag: "Substatus failed", detail: `Substatus not applied: ${err.message}` });
      }
    }
  }

  return { jobNumber, warnings };
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
async function emailHtmlForNote(html, oneDriveUrl = null, emailMeta = null) {
  let cleaned = html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<img[^>]*>/gi, "")
    .trim();

  // Remove Inky security banner — strip from start of HTML up to and including the
  // ipw-end anchor (raw email uses "ipw-end-*", Outlook renderer prefixes it "x_ipw-end-*")
  cleaned = cleaned.replace(/[\s\S]*?<a\b[^>]*name="(?:x_)?ipw-end-\d+"[^>]*><\/a>(?:\s*<\/\w+>)*/i, "");

  cleaned = await decodeWrappedLinks(cleaned);

  const cell = (label, value) =>
    `<tr><td style="color:#888888;font-size:12px;font-weight:bold;padding:1px 12px 1px 0;white-space:nowrap;vertical-align:top">${label}</td><td style="color:#444444;font-size:12px;padding:1px 0">${value}</td></tr>`;

  const metaRows = [
    emailMeta?.from    ? cell("From:",       emailMeta.from)    : "",
    emailMeta?.to      ? cell("To:",         emailMeta.to)      : "",
    emailMeta?.subject ? cell("Subject:",    emailMeta.subject) : "",
    oneDriveUrl        ? cell("Attachment:", `<a href="${oneDriveUrl}" style="color:#1a6bbf" target="_blank">View Work Order PDF</a>`) : "",
  ].filter(Boolean).join("");

  const titleRow = `<tr><td colspan="2" style="font-size:16px;font-weight:bold;color:#444444;padding:0 0 5px 0">Work Order</td></tr>`;
  const metaHtml = `<table style="border-collapse:collapse;margin:0 0 12px 0">${titleRow}${metaRows}</table>`;

  return `${metaHtml}<hr style="border:none;border-top:1px solid #dddddd;margin:0 0 14px 0"><div>${cleaned}</div>`;
}

// A plain driveItem webUrl opens the file in isolation with no folder context, so there's
// no gallery navigation. SharePoint's "browse in folder" deep link (onedrive.aspx with id +
// parent params) loads the file's parent folder alongside it, which gives the native
// prev/next arrow navigation between sibling files.
function buildFolderPreviewUrl(webUrl) {
  const u = new URL(webUrl);
  const filePath = decodeURIComponent(u.pathname);
  const parentPath = filePath.slice(0, filePath.lastIndexOf("/"));
  const segments = filePath.split("/").filter(Boolean); // ["sites", "BaraElectricalServices", ...]
  const siteBase = `${u.origin}/${segments[0]}/${segments[1]}`;
  return `${siteBase}/_layouts/15/onedrive.aspx?id=${encodeURIComponent(filePath)}&parent=${encodeURIComponent(parentPath)}`;
}

// Small clickable thumbnail grid, posted as its own note. Clicking a thumbnail opens that
// specific photo in SharePoint with its folder context loaded; because all the job's photos
// live in the same folder, the tech can still arrow through the rest from there.
function buildPhotoGalleryNote(photos) {
  const thumbs = photos.map(p =>
    `<a href="${buildFolderPreviewUrl(p.webUrl)}" target="_blank"><img src="${p.thumbnailUrl || p.webUrl}" alt="${p.name}" style="width:110px;height:110px;object-fit:cover;margin:4px;border-radius:4px;border:1px solid #dddddd" /></a>`
  ).join("");

  const titleRow = `Photos (${photos.length})`;
  return `<div style="font-size:14px;font-weight:bold;color:#444444;margin:0 0 8px 0">${titleRow}</div><div>${thumbs}</div>`;
}

// Cached drive ID for the Bara Electrical Services SharePoint document library
let sharepointDriveId = null;

async function getSharepointDriveId() {
  if (sharepointDriveId) return sharepointDriveId;
  const siteRes  = await graphFetch(`/sites/baraelectrical.sharepoint.com:/sites/BaraElectricalServices`);
  const site     = await siteRes.json();
  const drivesRes = await graphFetch(`/sites/${site.id}/drives`);
  const drives   = await drivesRes.json();
  const drive    = (drives.value || []).find(d => d.name === "Documents" || d.name === "Shared Documents");
  if (!drive) throw new Error("SharePoint Documents drive not found");
  sharepointDriveId = drive.id;
  return sharepointDriveId;
}

async function uploadWorkOrderToOneDrive(jobNumber, filename, contentBytes, contentType = "application/pdf") {
  const safeName = filename.replace(/#/g, "").trim();
  const itemPath = ["General", "Other", "AI Workorders [dont touch]", `${jobNumber} - ${safeName}`]
    .map(s => encodeURIComponent(s)).join("/");
  const uploadData = await putSharepointFile(itemPath, contentBytes, contentType);
  return uploadData.webUrl || null;
}

// Photos go in a per-job subfolder (rather than the flat filename-prefixed layout used for
// the PDF) so that opening any one of them from the note still gives SharePoint's native
// gallery view — arrow through the rest of the job's photos — instead of a dead-end preview.
async function uploadPhotoToOneDrive(jobNumber, filename, contentBytes, contentType) {
  const safeName = filename.replace(/#/g, "").trim();
  const itemPath = ["General", "Other", "AI Workorders [dont touch]", jobNumber, safeName]
    .map(s => encodeURIComponent(s)).join("/");
  const uploadData = await putSharepointFile(itemPath, contentBytes, contentType);
  return { id: uploadData.id, webUrl: uploadData.webUrl };
}

// Thumbnail image URL for a driveItem, to embed as a small clickable preview in the note.
// These Graph-issued URLs are short-lived (not the permanent webUrl) — fine for a tech
// checking the job soon after creation, but the thumbnail image itself can go stale later
// even though the click-through link keeps working.
async function getThumbnailUrl(driveId, itemId) {
  const token = await getAccessToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/thumbnails`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) return null;
  const data = await res.json();
  return data.value?.[0]?.medium?.url || data.value?.[0]?.large?.url || null;
}

async function putSharepointFile(itemPath, contentBytes, contentType) {
  const driveId = await getSharepointDriveId();

  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), 20000);

  try {
    const token = await getAccessToken();
    const uploadRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${itemPath}:/content`,
      { method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": contentType }, body: contentBytes, signal: ac.signal }
    );
    if (!uploadRes.ok) {
      const err = await uploadRes.json().catch(() => ({}));
      throw new Error(`SharePoint upload failed ${uploadRes.status}: ${err?.error?.message || JSON.stringify(err)}`);
    }
    return await uploadRes.json();
  } finally {
    clearTimeout(timer);
  }
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
    if (/work\s*order/i.test(text)) return href;
  }

  // Fallback: known portal domains or "workorder" in the URL itself
  for (const [, href] of anchors) {
    if (!href.startsWith("https://")) continue;
    const dest = /safelinks\.protection\.outlook\.com/i.test(href)
      ? (() => { try { return decodeURIComponent(new URL(href).searchParams.get("url") || href); } catch { return href; } })()
      : href;
    if (/workorder/i.test(dest) || WORKORDER_DOMAINS.test(dest)) return href;
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
async function processMessage(message, mailbox = WORKORDERS_EMAIL, onStatus = null) {

  const attachments = message.attachments || [];
  const rawBody     = message.body?.content || "";
  const emailBodyText = (message.body?.contentType === "html" ? cleanHtml(rawBody) : rawBody).replace(/\s+/g, " ").trim();
  let   textForAI     = emailBodyText;
  let   pdfAttachment = null;

  function withEmailBody(primary) {
    return `--- WORK ORDER CONTENT (prefer this) ---\n${primary}\n\n--- EMAIL BODY (use for anything not found above) ---\n${emailBodyText}`;
  }

  // Work order link detection
  const workOrderLink = findWorkOrderLink(rawBody);

  if (workOrderLink) {
    try {
      let response = await fetch(workOrderLink, {
        redirect: "follow",
        headers: { "User-Agent": "Mozilla/5.0" },
      });

      // Inky link protection stops the redirect — bypass with confirm=True
      if (response.url.includes("shared.outlook.inky.com") && !response.url.includes("confirm=True")) {
        response = await fetch(response.url + "&confirm=True", {
          redirect: "follow",
          headers: { "User-Agent": "Mozilla/5.0" },
        });
      }

      const contentType = response.headers.get("content-type") || "";
      if (contentType.includes("pdf")) {
        const buffer  = await response.arrayBuffer();
        const pdfText = await extractPDF(new Uint8Array(buffer));
        textForAI = withEmailBody(pdfText.replace(/\s+/g, " ").trim());
      } else {
        const html     = await response.text();
        const linkText = cleanHtml(html).slice(0, 50000);
        if (linkText.length > 200) textForAI = withEmailBody(linkText);
      }
    } catch (err) {
      console.error("Link fetch error:", err.message);
    }
  }

  // Work order PDF attachment — takes priority over link
  const workorderAttachment = attachments.find(a =>
    (a.name || "").toLowerCase().endsWith(".pdf")
  );

  if (workorderAttachment) {
    const attachRes  = await graphFetch(
      `/users/${mailbox}/messages/${message.id}/attachments/${workorderAttachment.id}`
    );
    const attachData = await attachRes.json();
    const data       = Uint8Array.from(atob(attachData.contentBytes), c => c.charCodeAt(0));
    pdfAttachment = { name: workorderAttachment.name, data: new Uint8Array(data) }; // copy before pdfjs detaches the buffer
    const pdfText    = await extractPDF(data);
    textForAI = withEmailBody(pdfText.replace(/\s+/g, " ").trim());
  }
  if (onStatus) await onStatus(SENDING_TO_AI_CATEGORY);

  const responseAI = await openai.responses.create({
    model: "gpt-5-mini",
    text: { format: { type: "json_object" } },
    instructions: `You are a work order extraction system for an electrical company in Australia. Today's date is ${new Date().toLocaleDateString("en-AU", { day: "2-digit", month: "2-digit", year: "numeric" })}.

CRITICAL RULES:
- tenant-name and tenant-contact come from the Tenant Details section, OR any section labelled "Contact for job access" or similar, OR an inline mention anywhere in the description/notes such as "contact tenant emma for access 0414152246" (extract "Emma" as tenant-name and "0414152246" as tenant-contact in that case). Always scan the full description/instructions text for a phrase like "contact tenant <name>" or "tenant <name> on <number>" even if there is no dedicated tenant section. Never use the Owner Details / Owner section as tenant-name, tenant-contact, or tenant-email — the owner is not the tenant, even if the PM states tenant details were not provided. If no tenant is listed anywhere (including inline) and the property is explicitly stated as vacant, set tenant-name to "Vacant" and ensure access-details captures any lockbox or key collection info. If no tenant is listed and the property is NOT stated as vacant, leave tenant-name null. Never use access details, lockbox info, or key numbers as the tenant name. If multiple tenants are listed, include ALL of them separated by commas — do not drop any.
- If the email states the tenant is vacating or moving out and the move-out date is within 7 days of today, treat the property as vacant: set tenant-name to "Vacant", set tenant-contact to null, and include in task-description that keys should be collected after the move-out date (e.g. "Collect keys after 3/7").
- access-details is ONLY physical access codes/numbers — key numbers, lockbox codes, gate codes, swipe card numbers. e.g. "Key: 1234", "Lockbox code: 56", "Gate code: 789". Do NOT include contact instructions, tenant names, safety instructions, or anything that is not a physical code or number.
- expenditure-limit is the dollar amount only — e.g. "$330". Strip any conditions, notes, or extra text after the amount. If the expenditure limit is $0 or zero, return null.
- confidence is a float 0.0–1.0 rating how confident you are in the overall extraction. 1.0 = all fields clearly present, 0.0 = guessing most fields.
- notes is any concerns, ambiguities, or flags worth mentioning — e.g. missing fields, conflicting info, unusual job details. Leave null if nothing to flag.
- tenant-contact must contain phone numbers ONLY — no names, no labels, just the numbers. Only use a number if it is explicitly and unambiguously tied to the tenant (e.g. appears in a Tenant section, is labelled "Tenant Phone"/"Tenant Mobile"/"Contact Number", or immediately follows an inline "contact tenant <name>" style phrase). If you are unsure whether a number belongs to the tenant, leave tenant-contact null. If there are multiple confirmed tenant numbers, separate with commas. Prefer mobile over home numbers. Australian numbers always start with 0 (e.g. 0412 345 678) — always include the leading 0.
- tenant-email is the tenant's email address. Only include if explicitly labelled as the tenant's email. Leave null if not present or uncertain.
- property-manager comes from the Property Manager section, OR from an Agency Details section where the manager is listed (e.g. "Manager: Jane Smith"). Use the person's name only, not the agency name.
- account-to must include ALL owners exactly as written, always in the format: owners c/o real estate.
- real-estate must always be a company or agency name — never a URL or domain. If the source contains something like "aussieproperty.com.au", convert it to a readable name (e.g. "Aussie Property") by stripping the domain extension and formatting as a proper name. If you cannot find it directly, look for it in account-to after the c/o. The sender's email address is provided at the top of the input — use the domain as an additional hint to identify real-estate if the company name is not clearly stated in the content (e.g. "noreply@raywhite.com.au" → "Ray White").
- order-number is the job/work order number.
- address is required — if it isn't clearly stated in the body/PDF content, check the email subject line (provided at the top of the input) since it often contains the property address.
- task-description must be a concise electrician job summary. If anything is listed as conditional or requires approval (e.g. "deluxe clean if approved", "AC2 if required"), include that in the description too.
- Do NOT include instructions to contact the tenant or PM for access in task-description. Contacting the tenant for access is the default assumption for every job and must not be stated.
- Key numbers in access-details are reference numbers for our existing key management system — we already hold these keys. Do NOT include any instruction to collect, pick up, or obtain keys in task-description based solely on a key number being listed in access-details.
- Do NOT guess missing fields — if missing return null.
- The text may be a structured form (with clear sections) OR plain prose in an email. Extract the same fields either way — don't return null just because sections aren't labelled.
- If the input contains both a PDF CONTENT section and an EMAIL BODY section, prefer the PDF for all fields but check the email body for anything not found in the PDF.
- For task-type: if the text explicitly mentions EC1, AC1, AC2, or ACEC1 anywhere, use that — it takes priority over everything else, even if other work is also mentioned. If a compliance check or aircon service is recommended or suggested (e.g. "recommend compliance check", "suggest an EC1"), treat it as that task type — recommendations from a PM or owner should be acted on. EXCEPTION: if the work order includes a package (EC1/AC1/AC2/ACEC1) AND additional work beyond the package, set task-type to "Real Estate General Maintenance" (or "Real Estate Aircon Maintenance" if the extra work is aircon-related) — but still set the package field to the package code. NOTE: smoke alarm testing, replacement, or supply is already included within EC1 — do not treat it as additional work beyond the package.
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
  "tenant-email": "",
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
    input: `From: ${message.from?.emailAddress?.address || ""}\nSubject: ${message.subject || ""}\n\nExtract the following work order and return JSON:\n\n${textForAI}`,
  });

  const parsed = JSON.parse(responseAI.output_text);

  // Normalise address to title case in case the source had ALL CAPS suburb
  if (parsed.address) {
    parsed.address = parsed.address.replace(/,?\s*Australia\s*$/i, "").trim();
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

  // Ignore $0 expenditure limits
  if (parsed["expenditure-limit"] && /^\$?\s*0+(\.0+)?\s*$/.test(parsed["expenditure-limit"].trim())) {
    parsed["expenditure-limit"] = null;
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

  const emailMeta = {
    from:    message.from?.emailAddress?.address || null,
    to:      (message.toRecipients || []).map(r => r.emailAddress?.address).filter(Boolean).join(", ") || null,
    subject: message.subject || null,
  };

  // Download image attachments in parallel
  const imageAttachments = (await Promise.all(
    attachments
      .filter(a => !/inky/i.test(a.name || ""))
      .filter(a => /\.(jpe?g|png|gif|bmp|webp)$/i.test(a.name || "") || (a.contentType || "").startsWith("image/"))
      .map(async a => {
        try {
          const attRes  = await graphFetch(`/users/${mailbox}/messages/${message.id}/attachments/${a.id}`);
          const attData = await attRes.json();
          const imgData = Uint8Array.from(atob(attData.contentBytes), c => c.charCodeAt(0));
          console.log("Image attachment downloaded:", a.name);
          return { name: a.name, data: imgData, contentType: a.contentType || "image/jpeg" };
        } catch (err) {
          console.warn("Failed to download image attachment:", a.name, err.message);
          return null;
        }
      })
  )).filter(Boolean);

  return { result: parsed, rawEmail: rawBody, pdfAttachment, imageAttachments, emailMeta };
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
        await logAiOutput(result, message.subject);

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

// Look for a sibling message already tagged "Job created - X" or "Existing job - X"
// elsewhere in the same reply thread — replies to an already-processed work
// order get their own "Bara AI" tag but shouldn't trigger their own job.
// Mailbox-wide, not Inbox-scoped — existing filing rules (e.g. "move mail from
// this client into its own subfolder") can relocate a message out of Inbox
// well before a later reply in the same thread gets processed.
async function findJobTagInThread(mailbox, conversationId, excludeMessageId) {
  const filter = encodeURIComponent(`conversationId eq '${conversationId}'`);
  const res = await graphFetch(
    `/users/${mailbox}/messages?$filter=${filter}&$select=id,categories&$top=25`
  );
  const data = await res.json();
  if (!res.ok) return null;
  for (const m of (data.value || [])) {
    if (m.id === excludeMessageId) continue;
    const tag = (m.categories || []).find(c => c.startsWith("Job created") || c.startsWith("Existing job"));
    if (tag) return tag;
  }
  return null;
}

// Stamp the same job tag onto every other message already in the thread, so
// anyone looking at any message in the conversation sees the job status
// immediately instead of just the "Bara AI" trigger tag.
async function tagWholeConversation(mailbox, conversationId, excludeMessageId, tag) {
  if (!conversationId) return;
  try {
    const filter = encodeURIComponent(`conversationId eq '${conversationId}'`);
    const res = await graphFetch(
      `/users/${mailbox}/messages?$filter=${filter}&$select=id,categories&$top=25`
    );
    const data = await res.json();
    if (!res.ok) return;
    for (const m of (data.value || [])) {
      if (m.id === excludeMessageId) continue;
      const categories = [
        ...(m.categories || []).filter(c => !STATUS_CATEGORIES.includes(c) && !c.startsWith("Job created") && !c.startsWith("Existing job")),
        tag,
      ];
      try {
        await graphFetch(`/users/${mailbox}/messages/${m.id}`, {
          method: "PATCH",
          body: JSON.stringify({ categories }),
        });
      } catch (err) {
        console.warn("Failed to tag conversation message:", m.id, err.message);
      }
    }
  } catch (err) {
    console.warn("tagWholeConversation failed:", err.message);
  }
}

async function pollInbox(mailbox) {
  const filter = encodeURIComponent(
    `categories/any(c:c eq '${TRIGGER_CATEGORY}')` +
    ` and not categories/any(c:c eq '${CLIENT_NOT_FOUND_CATEGORY}')` +
    ` and not categories/any(c:c eq '${NO_ADDRESS_CATEGORY}')` +
    ` and not categories/any(c:c eq '${PROCESSING_CATEGORY}')` +
    ` and not categories/any(c:c eq '${FAILED_CATEGORY}')` +
    ` and not categories/any(c:c eq '${READING_EMAIL_CATEGORY}')` +
    ` and not categories/any(c:c eq '${SENDING_TO_AI_CATEGORY}')` +
    ` and not categories/any(c:c eq '${CREATING_JOB_CATEGORY}')`
  );

  const res  = await graphFetch(
    `/users/${mailbox}/mailFolders/inbox/messages` +
    `?$filter=${filter}` +
    `&$select=id,subject,body,categories,from,toRecipients,conversationId,receivedDateTime` +
    `&$expand=attachments($select=id,name,contentType,size)` +
    `&$top=10`
  );
  const data = await res.json();
  if (!res.ok) throw new Error(`Graph API error ${res.status}: ${JSON.stringify(data?.error || data)}`);

  // Newest first — when a thread has multiple untagged messages, this ensures the
  // most recent one (which quotes everything underneath it) is the one processed
  // and posted to Aroflo; the earlier ones get caught as duplicates by the thread
  // check below since the newest is tagged first.
  const messages = (data.value || [])
    .filter(m => !m.categories.some(c => c.startsWith("Job created")))
    .sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
  console.log(`Poll (${mailbox}): ${messages.length} email(s) found`);

  for (const message of messages) {
    const siblingTag = message.conversationId
      ? await findJobTagInThread(mailbox, message.conversationId, message.id)
      : null;
    if (siblingTag) {
      const jobNumber = siblingTag.split(" - ")[1] || siblingTag;
      console.log(`Reply to already-processed thread — tagging as duplicate of ${siblingTag}:`, message.subject);
      await graphFetch(`/users/${mailbox}/messages/${message.id}`, {
        method: "PATCH",
        body: JSON.stringify({ categories: [...message.categories.filter(c => !STATUS_CATEGORIES.includes(c)), `Existing job - ${jobNumber}`] }),
      });
      continue;
    }

    let currentCategories = await setJobStatus(mailbox, message.id, message.categories, READING_EMAIL_CATEGORY);
    console.log("Reading:", message.subject);

    try {
      const onStatus = async (status) => {
        currentCategories = await setJobStatus(mailbox, message.id, currentCategories, status);
      };

      const { result, rawEmail, pdfAttachment, imageAttachments, emailMeta } = await processMessage(message, mailbox, onStatus);

      // Pre-job validation warnings — checked before Aroflo job creation
      const preWarnings = [];
      if (!result["property-manager"]) {
        preWarnings.push({ tag: "No PM in email", detail: "Property manager not listed in the work order" });
      }
      if (!result["tenant-name"] && !result["tenant-contact"]) {
        preWarnings.push({ tag: "No tenant info", detail: "No tenant name or contact in the work order, and property not stated as vacant" });
      }

      await onStatus(CREATING_JOB_CATEGORY);
      const { jobNumber, warnings: jobWarnings } = await createArofloJob(result, rawEmail, pdfAttachment, emailMeta, imageAttachments);
      logAiOutput(result, message.subject).catch(err => console.warn("logAiOutput:", err.message));
      logActivity("Job created", jobNumber).catch(err => console.warn("logActivity:", err.message));
      const allWarnings = [...preWarnings, ...jobWarnings];

      // Always apply job tag to prevent re-processing; add a specific tag for each failure
      const jobTag = `Job created - ${jobNumber}`;
      const finalCategories = [
        ...currentCategories.filter(c => !STATUS_CATEGORIES.includes(c)),
        jobTag,
        ...allWarnings.map(w => w.tag),
      ];
      await graphFetch(`/users/${mailbox}/messages/${message.id}`, {
        method: "PATCH",
        body: JSON.stringify({ categories: finalCategories }),
      });
      currentCategories = finalCategories;
      tagWholeConversation(mailbox, message.conversationId, message.id, jobTag)
        .catch(err => console.warn("tagWholeConversation:", err.message));

      if (allWarnings.length > 0) {
        console.warn("Job created with issues:", allWarnings.map(w => w.tag));
        try {
          const warningLines = allWarnings.map(w => `<li style="margin:4px 0;font-family:sans-serif;font-size:14px"><strong>${w.tag}:</strong> ${w.detail}</li>`).join("");
          await graphFetch(`/users/${WORKORDERS_EMAIL}/sendMail`, {
            method: "POST",
            body: JSON.stringify({
              message: {
                subject: `Action required — Job ${jobNumber} created with issues`,
                toRecipients: [{ emailAddress: { address: BRANDON_EMAIL } }],
                body: {
                  contentType: "HTML",
                  content: `<p style="font-family:sans-serif;font-size:14px">Job <strong>${jobNumber}</strong> was created in Aroflo but the following need attention:</p><ul>${warningLines}</ul><p style="font-family:sans-serif;font-size:12px;color:#888">Original email: ${message.subject}</p>`,
                },
              },
              saveToSentItems: false,
            }),
          });
          console.log("Alert email sent for job:", jobNumber);
        } catch (alertErr) {
          console.error("Failed to send alert email for job", jobNumber, alertErr.message);
        }
      } else {
        console.log("Tagged as done:", message.subject);
      }
    } catch (err) {
      console.error("Error processing message:", message.subject, err.message);
      if (err.message.startsWith("Client not found")) {
        await setJobStatus(mailbox, message.id, currentCategories, CLIENT_NOT_FOUND_CATEGORY);
        console.log("Tagged as client not found:", message.subject);
      } else if (err.message.startsWith("No address found")) {
        await setJobStatus(mailbox, message.id, currentCategories, NO_ADDRESS_CATEGORY);
        console.log("Tagged as no address:", message.subject);
      } else {
        // Tag as "Failed" — remove it in Outlook to retry
        await setJobStatus(mailbox, message.id, currentCategories, FAILED_CATEGORY);
        console.log("Tagged as failed:", message.subject);
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
    await pollInbox(WORKORDERS_EMAIL);
    await pollInbox(BRANDON_EMAIL);
  } catch (err) {
    console.error("Poll error:", err);
  } finally {
    pollRunning = false;
  }
}


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
// TEMP: Dump client cache as JSON — GET /clients
// ================================================================
app.get("/clients", (req, res) => {
  res.json({
    count: clientCache.size,
    clients: [...clientCache.values()].map(c => ({ clientid: c.clientid, clientname: c.clientname })),
  });
});

app.get("/find-client", async (req, res) => {
  const name = req.query.name;
  if (!name) return res.status(400).json({ error: "Pass ?name=..." });
  const result = await findClient(name);
  const key = name.toLowerCase();
  const exactCacheHit = clientCache.get(key);
  const partialMatches = [...clientCache.values()]
    .filter(c => c.clientname.toLowerCase().includes(key) || key.includes(c.clientname.toLowerCase()))
    .map(c => c.clientname);
  res.json({ result, exactCacheHit: exactCacheHit?.clientname || null, partialMatches });
});

// TEMP: exercise the per-job photo-folder + thumbnail-gallery note flow against a real job,
// without going through the full email pipeline. GET /test-photo-note?job=103681
// ================================================================
app.get("/test-photo-note", async (req, res) => {
  const jobNumber = req.query.job;
  if (!jobNumber) return res.status(400).json({ error: "Pass ?job=<jobnumber>" });
  try {
    const fetched = await arofloGet("zone=tasks&where=" + encodeURIComponent(`and|jobnumber|=|${jobNumber}`) + "&page=1");
    const arr = Array.isArray(fetched.tasks) ? fetched.tasks : [fetched.tasks];
    const taskId = arr[0]?.taskid;
    if (!taskId) return res.json({ error: `No task found for job ${jobNumber}` });

    // Small solid-color PNG test fixtures — real enough for SharePoint to generate thumbnails.
    const redPng = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAAAI0lEQVR4nO3BAQ0AAADCoPdPbQ43oAAAAAAAAAAAAAAAAHgaJAAAAdE2QOsAAAAASUVORK5CYII=",
      "base64"
    );
    const bluePng = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAAAI0lEQVR4nO3BMQEAAADCIPunfjkKYAAAAAAAAAAAAAAAAOA1JAAAAV8Q3TQAAAAASUVORK5CYII=",
      "base64"
    );
    const driveId = await getSharepointDriveId();
    const item1 = await uploadPhotoToOneDrive(jobNumber, "test-photo-1.png", redPng, "image/png");
    const item2 = await uploadPhotoToOneDrive(jobNumber, "test-photo-2.png", bluePng, "image/png");
    const photos = [
      { name: "test-photo-1.png", webUrl: item1.webUrl, thumbnailUrl: await getThumbnailUrl(driveId, item1.id).catch(() => null) },
      { name: "test-photo-2.png", webUrl: item2.webUrl, thumbnailUrl: await getThumbnailUrl(driveId, item2.id).catch(() => null) },
    ];

    const noteHtml = `<p><strong>Test note</strong> — photo gallery check.</p>` + buildPhotoGalleryNote(photos);
    const updateXml =
`<tasks>
  <task>
    <taskid>${taskId}</taskid>
    <notes><note><content><![CDATA[${noteHtml}]]></content></note></notes>
  </task>
</tasks>`;
    const upZone = await arofloPost("zone=tasks&postxml=" + encodeURIComponent(updateXml));

    res.json({
      taskId,
      photos: photos.map(p => ({ ...p, galleryUrl: buildFolderPreviewUrl(p.webUrl) })),
      noteApplied: Number(upZone.postresults?.updatetotal ?? 0) > 0,
    });
  } catch (err) {
    res.json({ error: err.message });
  }
});

// ================================================================
// AROFLO WEBHOOK — new client created
// ================================================================
app.post("/aroflo-webhook", express.json(), async (req, res) => {
  console.log("Aroflo webhook received:", JSON.stringify(req.body));
  res.sendStatus(200);
  await loadClientCache();
});

// ================================================================
// START
// ================================================================
app.listen(process.env.PORT || 3000, async () => {
  console.log(`Server running — deployed ${new Date().toISOString()}`);
  // Wait for the client cache before polling — otherwise a work order sitting
  // in the inbox at deploy time gets processed against an empty cache and
  // falls back to a less reliable single-candidate live API lookup.
  await loadClientCache();
  pollEmails();
  forwardRicaEmails();
  processAiTestingEmails();
  setInterval(pollEmails, POLL_INTERVAL_MS);
  setInterval(forwardRicaEmails, POLL_INTERVAL_MS);
  setInterval(processAiTestingEmails, POLL_INTERVAL_MS);
});

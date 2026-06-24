import express from "express";
import OpenAI from "openai";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

const REQUIRED_ENV = ["OPENAI_API_KEY", "GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET", "GRAPH_RECIPIENT"];
for (const key of REQUIRED_ENV) {
  if (!process.env[key]) {
    console.error(`${key} is not set`);
    process.exit(1);
  }
}

const TRIGGER_CATEGORY = "Bara AI";
const DONE_CATEGORY = "Job created";
const POLL_INTERVAL_MS = 60 * 1000;

const app = express();
app.use(express.json({ limit: "25mb" }));

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// ---------------- Graph API auth ----------------
let tokenCache = { token: null, expiry: 0 };

async function getAccessToken() {
  if (tokenCache.token && Date.now() < tokenCache.expiry - 60000) {
    return tokenCache.token;
  }
  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: process.env.GRAPH_CLIENT_ID,
        client_secret: process.env.GRAPH_CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default",
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error(`Auth failed: ${JSON.stringify(data)}`);
  tokenCache = { token: data.access_token, expiry: Date.now() + data.expires_in * 1000 };
  return tokenCache.token;
}

async function graphFetch(path, options = {}) {
  const token = await getAccessToken();
  return fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });
}

// ---------------- PDF extraction ----------------
async function extractPDF(data) {
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;
  let text = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(item => item.str).join(" ") + "\n";
  }
  return text;
}

// ---------------- HTML cleaning ----------------
function cleanHtml(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// ---------------- Process a single message ----------------
async function processMessage(message) {
  console.log("WORK ORDER EMAIL RECEIVED:", message.subject);

  const attachments = message.attachments || [];
  let textForAI = message.body?.contentType === "html"
    ? cleanHtml(message.body.content)
    : (message.body?.content || "");

  // ---------------- TAPI LINK DETECTION ----------------
  const tapiMatch = textForAI.match(/https:\/\/url\d+\.tapihq\.com\/ls\/click\S+/i);

  if (tapiMatch) {
    console.log("RAW TAPI MATCH:", tapiMatch[0]);

    const tapiLink = tapiMatch[0]
      .split(">")[0]
      .split('"')[0]
      .split(")")[0]
      .trim();

    console.log("CLEAN TAPI LINK:", tapiLink);

    try {
      const response = await fetch(tapiLink, {
        redirect: "follow",
        headers: { "User-Agent": "Mozilla/5.0" },
      });

      console.log("FINAL URL:", response.url);

      const html = await response.text();
      console.log("RAW HTML LENGTH:", html.length);
      console.log("HTML PREVIEW:", html.slice(0, 300));

      textForAI = cleanHtml(html).slice(0, 50000);
      console.log("CLEAN TEXT LENGTH:", textForAI.length);
    } catch (err) {
      console.error("TAPI PROCESSING ERROR:", err);
    }
  }

  // ---------------- WORK ORDER PDF FLOW ----------------
  const workorderAttachment = attachments.find(a => {
    const name = (a.name || "").toLowerCase();
    return name.includes("workorder") && name.endsWith(".pdf");
  });

  if (workorderAttachment) {
    console.log("FOUND WORK ORDER PDF:", workorderAttachment.name);

    const attachRes = await graphFetch(
      `/users/${process.env.GRAPH_RECIPIENT}/messages/${message.id}/attachments/${workorderAttachment.id}`
    );
    const attachData = await attachRes.json();
    const data = Uint8Array.from(atob(attachData.contentBytes), c => c.charCodeAt(0));
    const pdfText = await extractPDF(data);

    textForAI = pdfText.replace(/\s+/g, " ").trim();
    console.log("USING PDF CONTENT FOR AI");
  } else if (!tapiMatch) {
    console.log("NO PDF OR TAPI - USING EMAIL BODY");
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

// ---------------- Poll for tagged emails ----------------
async function pollEmails() {
  try {
    const filter = encodeURIComponent(
      `categories/any(c:c eq '${TRIGGER_CATEGORY}') and not categories/any(c:c eq '${DONE_CATEGORY}')`
    );

    const res = await graphFetch(
      `/users/${process.env.GRAPH_RECIPIENT}/mailFolders/inbox/messages` +
      `?$filter=${filter}` +
      `&$select=id,subject,body,categories` +
      `&$expand=attachments($select=id,name,contentType,size)` +
      `&$top=10`
    );
    const data = await res.json();
    const messages = data.value || [];

    if (messages.length) console.log(`Found ${messages.length} email(s) to process`);

    for (const message of messages) {
      try {
        const result = await processMessage(message);
        console.log("AI RESULT:", result);

        // Tag as done
        await graphFetch(`/users/${process.env.GRAPH_RECIPIENT}/messages/${message.id}`, {
          method: "PATCH",
          body: JSON.stringify({ categories: [...message.categories, DONE_CATEGORY] }),
        });
        console.log("Tagged as done:", message.subject);
      } catch (err) {
        console.error("Error processing message:", message.subject, err);
      }
    }
  } catch (err) {
    console.error("Poll error:", err);
  }
}

// ---------------- Start server ----------------
app.listen(process.env.PORT || 3000, () => {
  console.log("Server running");
  pollEmails();
  setInterval(pollEmails, POLL_INTERVAL_MS);
});

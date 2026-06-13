import express from "express";
import bodyParser from "body-parser";
import OpenAI from "openai";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

const app = express();
app.use(bodyParser.json({ limit: "25mb" }));

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

// ---------------- PDF extraction (attachments only) ----------------
async function extractPDF(data) {
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;

  let text = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();

    const strings = content.items.map(item => item.str);
    text += strings.join(" ") + "\n";
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

// ---------------- EMAIL ROUTE ----------------
app.post("/email", async (req, res) => {
  try {
    console.log("WORK ORDER EMAIL RECEIVED");

    const attachments = req.body.attachments || [];
    let textForAI = req.body.text || "";

    // ---------------- TAPI LINK DETECTION ----------------
    const emailText = req.body.text || "";

    const tapiMatch = emailText.match(
      /https:\/\/url\d+\.tapihq\.com\/ls\/click\S+/i
    );

    if (tapiMatch) {
      console.log("RAW TAPI MATCH:", tapiMatch[0]);

      // 🔥 FIX: clean broken URL (removes > " ) etc)
      let tapiLink = tapiMatch[0]
        .split(">")[0]
        .split('"')[0]
        .split(")")[0]
        .trim();

      console.log("CLEAN TAPI LINK:", tapiLink);

      try {
        const redirectResponse = await fetch(tapiLink, {
          redirect: "follow",
          headers: {
            "User-Agent": "Mozilla/5.0"
          }
        });

        const finalUrl = redirectResponse.url;
        console.log("FINAL URL:", finalUrl);

        const finalPage = await fetch(finalUrl, {
          headers: {
            "User-Agent": "Mozilla/5.0"
          }
        });

        const html = await finalPage.text();

        console.log("RAW HTML LENGTH:", html.length);
        console.log("HTML PREVIEW:", html.slice(0, 300));

        const pageText = cleanHtml(html);

        console.log("CLEAN TEXT LENGTH:", pageText.length);

        textForAI = pageText;

      } catch (err) {
        console.error("TAPI PROCESSING ERROR:", err);
      }
    }

    // ---------------- WORK ORDER PDF FLOW ----------------
    const workorder = attachments.find(a => {
      const name = (a.filename || "").toLowerCase();
      return name.includes("workorder") && name.endsWith(".pdf");
    });

    if (workorder?.contentUrl) {
      console.log("FOUND WORK ORDER PDF:", workorder.filename);

      const response = await fetch(workorder.contentUrl);
      const arrayBuffer = await response.arrayBuffer();

      const data = new Uint8Array(arrayBuffer);

      const pdfText = await extractPDF(data);

      textForAI = pdfText.replace(/\s+/g, " ").trim();

      console.log("USING PDF CONTENT FOR AI");

    } else if (!tapiMatch) {
      console.log("NO PDF OR TAPI - USING EMAIL BODY");
      textForAI = (textForAI || "").replace(/\s+/g, " ").trim();
    }

    console.log("TEXT LENGTH:", textForAI.length);
    console.log("ABOUT TO CALL AI");

    const responseAI = await openai.responses.create({
      model: "gpt-4o-mini",

      text: {
        format: {
          type: "json_object"
        }
      },

      input: `
You are a work order extraction system for an electrical company.

CRITICAL RULES:
- tenant-name must ONLY come from Tenant Details section.
- property-manager must ONLY come from Property Manager section
- account-to must include ALL owners exactly as written
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

    console.log("AI RESULT:");
    console.log(responseAI.output_text);

    res.status(200).send("ok");

  } catch (err) {
    console.error("AI ERROR:", err);
    res.status(500).send("error");
  }
});

// ---------------- START SERVER ----------------
app.listen(process.env.PORT || 3000, () => {
  console.log("Server running");
});
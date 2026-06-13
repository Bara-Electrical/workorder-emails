import express from "express";
import bodyParser from "body-parser";
import OpenAI from "openai";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { chromium } from "playwright";

const app = express();
app.use(bodyParser.json({ limit: "25mb" }));

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

// ---------------- PDF extraction ----------------
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

// ---------------- Generate PDF from webpage ----------------
async function generatePdfFromUrl(url) {
  console.log("OPENING TAPI PAGE:", url);

  const browser = await chromium.launch({
    headless: true,
  });

  try {
    const page = await browser.newPage();

    await page.goto(url, {
      waitUntil: "networkidle",
      timeout: 60000,
    });

    const pdfBuffer = await page.pdf({
      format: "A4",
      printBackground: true,
    });

    return pdfBuffer;
  } finally {
    await browser.close();
  }
}

// ---------------- EMAIL ROUTE ----------------
app.post("/email", async (req, res) => {
  try {
    console.log("WORK ORDER EMAIL RECEIVED");

    const attachments = req.body.attachments || [];
    let textForAI = req.body.text || "";

    // ---------------- TAPI DETECTION ----------------
    const emailText = req.body.text || "";

    const tapiMatch = emailText.match(
      /https:\/\/url\d+\.tapihq\.com\/ls\/click\S+/i
    );

    if (tapiMatch) {
      console.log("FOUND TAPI LINK:", tapiMatch[0]);

      try {
        const redirectResponse = await fetch(tapiMatch[0], {
          redirect: "follow",
        });

        const finalUrl = redirectResponse.url;

        console.log("FINAL URL:", finalUrl);

        const pdfBuffer = await generatePdfFromUrl(finalUrl);

        const pdfData = new Uint8Array(pdfBuffer);

        let pdfText = await extractPDF(pdfData);

        textForAI = pdfText.replace(/\s+/g, " ").trim();

        console.log("USING GENERATED TAPI PDF FOR AI");
      } catch (err) {
        console.error("TAPI PROCESSING ERROR:", err);
      }
    }

    // ---------------- NORMAL PDF ATTACHMENT FLOW ----------------
    else {
      const workorder = attachments.find(a => {
        const name = (a.filename || "").toLowerCase();
        return name.includes("workorder") && name.endsWith(".pdf");
      });

      if (workorder?.contentUrl) {
        console.log("FOUND WORK ORDER PDF:", workorder.filename);

        const response = await fetch(workorder.contentUrl);
        const arrayBuffer = await response.arrayBuffer();

        const data = new Uint8Array(arrayBuffer);

        let pdfText = await extractPDF(data);

        textForAI = pdfText.replace(/\s+/g, " ").trim();

        console.log("USING PDF CONTENT FOR AI");
      } else {
        console.log("NO WORK ORDER PDF - USING EMAIL BODY");
        textForAI = (textForAI || "").replace(/\s+/g, " ").trim();
      }
    }

    console.log("ABOUT TO CALL AI");

    const responseAI = await openai.responses.create({
      model: "gpt-4o-mini",

      input: `
You are a work order extraction system for an electrical company.

CRITICAL RULES:
- tenant-name must ONLY come from Tenant Details section.
- if there is no tenant details the property may be vacant.
- if access details include a lockbox, put the lockbox details into tenant-name.
- property-manager must ONLY come from Property Manager section.
- account-to must include ALL owners exactly as written.
- do NOT guess missing fields.
- if missing return null.
- task-description must be concise electrician job summary.
- order-number is the job/work order number.

TASK TYPE RULES:
EC1 = Electrical Compliance Check (smoke alarms, RCDs, safety checks)
AC1 = Aircon Servicing
AC2 = Deluxe Aircon Clean
Real Estate Aircon Maintenance = anything involving air conditioning
Real Estate General Maintenance = everything else

Return JSON ONLY:

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
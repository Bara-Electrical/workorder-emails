import express from "express";
import bodyParser from "body-parser";
import OpenAI from "openai";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

const app = express();
app.use(bodyParser.json());

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

// ---------------- EMAIL ROUTE ----------------
app.post("/email", async (req, res) => {
  try {
    console.log("WORK ORDER EMAIL RECEIVED");

    const attachments = req.body.attachments || [];
    let textForAI = req.body.text || "";

    // Find workorder PDF
    const workorder = attachments.find(a => {
      const name = (a.filename || "").toLowerCase();
      return name.includes("workorder") && name.endsWith(".pdf");
    });

    if (workorder?.contentUrl) {
      console.log("FOUND WORK ORDER PDF:", workorder.filename);

      const response = await fetch(workorder.contentUrl);
      const arrayBuffer = await response.arrayBuffer();

      // REQUIRED: Uint8Array for pdfjs
      const data = new Uint8Array(arrayBuffer);

      let pdfText = await extractPDF(data);

      // Clean up PDF noise for better AI accuracy
      textForAI = pdfText.replace(/\s+/g, " ").trim();

      console.log("USING PDF CONTENT FOR AI");
    } else {
      console.log("NO WORK ORDER PDF - USING EMAIL BODY");
      textForAI = (textForAI || "").replace(/\s+/g, " ").trim();
    }

    console.log("ABOUT TO CALL AI");

    const responseAI = await openai.responses.create({
      model: "gpt-4o-mini",

      // NEW API FIX
      text: {
        format: {
          type: "json_object"
        }
      },

      input: `
You are a work order extraction system for an electrical company.

CRITICAL RULES:
- tenant-name must ONLY come from Tenant Details section
- if there is no tenant details the property might be vacant, it should have access details like a lockbox. add the lockbox into the tenant name and include the location and key.
- property-manager must ONLY come from Property Manager section
- account-to must include ALL owners exactly as written
- do NOT guess missing fields
- if missing return null
- task-description must be concise electrician job summary
- order-number is the job/work order number

TASK TYPE RULES:
EC1 = Electrical Compliance Check (smoke alarms, RCDs, safety checks)
AC1 = Aircon Servicing
AC2 = Deluxe Aircon Clean
Real Estate Aircon Maintenance = anything else that has aircon in the description, even if other works are involved aswell. other aircon jobs
Real Estate General Maintenance = everything else

Return ONLY valid JSON:

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
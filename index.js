import express from "express";
import bodyParser from "body-parser";
import OpenAI from "openai";
import { createRequire } from "module";

const app = express();
app.use(bodyParser.json());

const require = createRequire(import.meta.url);

// 🔥 IMPORTANT: direct internal import avoids ESM/CJS issues
const pdfParse = require("pdf-parse/lib/pdf-parse.js");

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

app.post("/email", async (req, res) => {
  try {
    console.log("WORK ORDER EMAIL RECEIVED");

    console.log(JSON.stringify(req.body, null, 2));

    const attachments = req.body.attachments || [];
    let textForAI = req.body.text || "";

    // Find workorder PDF (any variation)
    const workorder = attachments.find(a => {
      const name = (a.filename || "").toLowerCase();
      return name.includes("workorder") && name.endsWith(".pdf");
    });

    // If PDF exists, download + parse it
    if (workorder?.contentUrl) {
      console.log("FOUND WORK ORDER PDF:", workorder.filename);

      const response = await fetch(workorder.contentUrl);
      const arrayBuffer = await response.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);

      const pdfData = await pdfParse(buffer);

      textForAI = pdfData.text;

      console.log("USING PDF CONTENT FOR AI");
    } else {
      console.log("NO WORK ORDER PDF - USING EMAIL BODY");
    }

    console.log("ABOUT TO CALL AI");

    const responseAI = await openai.responses.create({
      model: "gpt-4o-mini",
      input: `
Extract details of a workorder from this text.

CRITICAL RULES:
- tenant-name ONLY from Tenant Details section
- property-manager ONLY from Property Manager section
- account-to must include all owners exactly as written
- if missing return null
- do not guess roles

Return JSON ONLY:
- task-type
- tenant-name
- tenant-contact
- address
- task-description
- real-estate
- property-manager
- account-to
- order-number

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

app.listen(process.env.PORT || 3000, () => {
  console.log("Server running");
});
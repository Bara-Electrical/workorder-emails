import express from "express";
import bodyParser from "body-parser";
import OpenAI from "openai";

const app = express();

app.use(bodyParser.json());

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

app.post("/email", async (req, res) => {
  try {
    console.log("WORK ORDER EMAIL RECEIVED");
    console.log("BODY:", req.body);

    const text = req.body.text || "";

    console.log("ABOUT TO CALL AI");

    const response = await openai.responses.create({
      model: "gpt-4o-mini",
      input: `
Extract details of a workorder from this email.
Rules:
- tenant-name must come ONLY from the "Tenant Details" section.
- property-manager must come ONLY from the "Property Manager Details" section.
- maintenance_request_posted_by is a separate field.
- Do not use names found elsewhere in the email for tenant-names.
- If multiple tenants exist, return all tenant names and contacts with commas between.
- If a section is missing, return null.
- The account-to is generally the owners name or multiple names followed by C/O and the real estate name. Make sure to include all owners names. This has to appear exactly as written.
- task-description is for an electricians job description no fluff. just list each fault with specific info.
- order-number will be a workorder number. It will say job number or similar.

Possible categories for task-type:

EC1 = Electrical Compliance Check - keywords= smoke alarm check, rcd check, ec1
AC1 = Aircon Servicing - keywords= ac1, aircon clean, aircon service
AC2 = Deluxe Aircon Clean - keywords= deluxe clean, deluxe service
Real Estate Aircon Maintenance = anything else to do with aircons
Real Estate General Maintenance = everything else if non of the above apply

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

Email:
${text}
      `,
    });

    console.log("AI RESULT:");
    console.log(response.output_text);

    res.status(200).send("ok");

  } catch (err) {
    console.error("AI ERROR:", err);
    res.status(500).send("error");
  }
});

app.listen(process.env.PORT || 3000, () => {
  console.log("Server running");
});
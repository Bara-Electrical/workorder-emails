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
Extract a work order from this email.

Return JSON ONLY:
- title
- address
- priority
- job_type
- notes

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
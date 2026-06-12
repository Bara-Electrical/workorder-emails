import express from "express";
import bodyParser from "body-parser";

const app = express();
app.use(bodyParser.json());

app.post("/email", (req, res) => {
  console.log("WORK ORDER EMAIL RECEIVED");

  console.log("Subject:", req.body.subject);
  console.log("From:", req.body.from);
  console.log("Body:", req.body.text);

  res.status(200).send("ok");
});

app.listen(process.env.PORT || 3000, () => {
  console.log("Workorder service running");
});
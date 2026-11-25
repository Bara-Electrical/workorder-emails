import assert from "node:assert/strict";
import { clientFromWebhook } from "../clientWebhook.js";

// Shape AroFlo uses for its other webhooks (compliance form, 23 Nov 2026).
assert.deepEqual(
  clientFromWebhook({ eventname: "Add Client", timestamputc: "2026-11-26 01:00:00", payload: { api_related_ids: { clientid: "JSc6UyNQXCAgCg==" }, clientname: "Acme Realty", currentuser: "Office User" } }),
  { clientId: "JSc6UyNQXCAgCg==", clientName: "Acme Realty" },
);
// Id inside payload rather than api_related_ids.
assert.deepEqual(clientFromWebhook({ payload: { clientid: "X1", clientname: "B" } }), { clientId: "X1", clientName: "B" });
// Flat body.
assert.deepEqual(clientFromWebhook({ clientid: "X2", clientname: "C" }), { clientId: "X2", clientName: "C" });
// Name only.
assert.deepEqual(clientFromWebhook({ payload: { clientname: "  D Realty " } }), { clientId: null, clientName: "D Realty" });
// Nothing usable.
assert.deepEqual(clientFromWebhook({}), { clientId: null, clientName: null });
assert.deepEqual(clientFromWebhook(null), { clientId: null, clientName: null });
assert.deepEqual(clientFromWebhook("Client created"), { clientId: null, clientName: null });

console.log("client-webhook tests passed");

# workorder-emails

Turns work-order emails in `workorders@baraelectrical.com.au` into AroFlo jobs.

Runs on Railway (project **Bara AI**, service **Workorder Emails**), deployed on push to
`main`. One process: `node index.js`.

## What it does

Every 2.5 minutes it reads the Inbox for emails tagged **Bara AI** (Outlook category) that
have not been dealt with yet, and for each one:

1. Reads the email, the work-order PDF (attached, linked, or on an earlier message in the
   thread) and any real photos in it. gpt-5-mini extracts the job details.
2. Finds the AroFlo client (name maps in `CLIENT_NAME_MAP` / `EMAIL_DOMAIN_MAP`), the
   location and the property manager contact, creating or updating the location as needed.
3. Asks the dashboard what it knows about that property (`gate.js`): jobs there in the last
   30 days, and the aircon units from compliance forms. The units go into the task
   description as a `Site: 2× Split` line.
4. Creates the task through the v1 API.
5. Puts the work order on the job the way AroFlo itself would (`aroflo-office.js`): the
   original email is forwarded to the task's own inbound address, and the PDF plus photos
   are uploaded to Documents & Photos through the v2 API.
6. Tags the email `Job created - <number>` (plus a tag per warning) and, if anything went
   wrong along the way, emails Brandon.

Replies in a thread that already has a job are tagged `Existing job - <number>` and never
create another one.

### The gate

`WORKORDER_GATE` decides whether a job at a property with a recent job is created straight
away or held for a person:

| value | behaviour |
|---|---|
| `off` (default) | checks run and are recorded; jobs always created |
| `tagged` | emails carrying the category **Gate Test** are held; the rest as `off` |
| `on` | every email with a recent job at its property is held |

A held email gets the category **Needs Decision**, no job, and a record on the dashboard.
The Bara Plugin shows the question under the email in Outlook with two answers; the
answer reaches this service through the dashboard's pending list on the next poll:
**New job — create it** re-runs the email with the gate skipped (category `Gate: Continue`),
**Same job — don't create** tags it **Stopped** for good. Every other message in that thread
is parked on the same category until the answer lands.

Design: `dashboard/docs/superpowers/specs/2026-09-12-workorder-gate-design.md`.

### The AroFlo office login

AroFlo only mints a task's inbound email address from a button in the office UI; neither
API can. So `aroflo-office.js` logs into office.aroflo.com as a service user and posts the
same RPC the button does. That user has MFA forced on it, so the service is its own
authenticator: it computes the TOTP code from `AROFLO_OFFICE_TOTP_SECRET`. It also clears
AroFlo's "User Session Limit" step, which appears when old sessions (earlier deploys) are
still alive. One login per process; it logs in again only when a request is bounced.

AroFlo may still refuse a login from a new IP with "Unusual login activity detected" — the
approval goes to the service user's mailbox. Approve it there once; the job is created
regardless, only the email/upload step is missed and the alert email says so.

## Environment

Required (the process refuses to start without them):

| variable | purpose |
|---|---|
| `OPENAI_API_KEY` | extraction and the photo filter |
| `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET` | Microsoft Graph, application permissions on the mailbox |
| `UENCODED`, `PENCODED`, `ORGENCODED`, `SECRET_KEY` | AroFlo v1 API (HMAC) |
| `AROFLO_V2_TOKEN` | AroFlo v2 API bearer token |
| `AROFLO_OFFICE_USER`, `AROFLO_OFFICE_PASS` | the service user's office.aroflo.com login |

Optional:

| variable | purpose |
|---|---|
| `AROFLO_OFFICE_TOTP_SECRET` | authenticator secret for the service user; without it MFA fails |
| `DASHBOARD_URL`, `DASHBOARD_API_SECRET` | the property check and gate records; without them jobs are created with a "Property check unavailable" warning |
| `WORKORDER_GATE` | `off` / `tagged` / `on`, see above |
| `ADMIN_API_KEY` | protects `/clients`, `/find-client`, `/aroflo-webhook`; unset means open |
| `AIRTABLE_API_KEY`, `AIRTABLE_BASE_ID` | activity log, best effort |

On Railway, `DASHBOARD_URL` and `DASHBOARD_API_SECRET` are references to the Chrome Plugin
service's variables of the same name.

### Rotating the authenticator

```
railway run node enrol-authenticator.mjs
```

Enrols a fresh secret on the service user and stores it on the Railway service. The secret
is never printed. To log in as that user in a browser, read the Railway variable into an
authenticator app.

## Tests

```
npm test
```

`node --test` for `aroflo-office.test.js` and `gate.test.js` (fake `fetch`, no network), then
the marker tests under `tools/`, which lift blocks of `index.js` by comment markers and run
them in isolation — keep those markers when editing.

## Local runs

`.env` holds the same variables for a local run of `node index.js`. To run a one-off script
with the Railway service's variables without copying secrets locally:

```
railway run node script.mjs
```

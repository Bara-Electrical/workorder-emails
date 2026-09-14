// One-off / rotation: enrol a fresh authenticator secret on the AroFlo service user and
// store it as AROFLO_OFFICE_TOTP_SECRET on the Railway service. Run from this folder:
//
//   railway run node enrol-authenticator.mjs
//
// `railway run` supplies AROFLO_OFFICE_USER/PASS; the secret is never printed.
import { createOfficeSession } from "./aroflo-office.js";
import { spawnSync } from "node:child_process";

const session = createOfficeSession();
try {
  await session.enrolAuthenticator(async (secret) => {
    // The CLI prints nothing on success; the exit code is the confirmation.
    const r = spawnSync(`railway variables --set "AROFLO_OFFICE_TOTP_SECRET=${secret}" --skip-deploys`,
      { encoding: "utf8", shell: true, stdio: ["ignore", "pipe", "pipe"] });
    if (r.status !== 0) throw new Error(`railway variables --set failed: ${`${r.stdout}${r.stderr}`.replaceAll(secret, "***").slice(0, 300)}`);
    console.log("Secret stored on Railway.");
  });
  console.log("Authenticator enrolled.");
} catch (err) {
  console.log("FAILED:", String(err.message).slice(0, 300));
  process.exitCode = 1;
}

// api/submit.js
import { google } from "googleapis";

// Map formType → Google Sheets tab name
const TAB_MAP = {
  "signup":            "Signups",
  "contact":           "Contact",
  "volunteer":         "Volunteers",
  "investor-interest": "Investor Interest",
  "investor-register": "Investor Register",
  "land-trust":        "Land Trust",
};

// Map formType → ordered row columns (each entry is [header, value])
function buildRow(formType, body) {
  const ts = new Date().toISOString();

  switch (formType) {
    case "signup":
      return [
        ts,
        body.firstName || "",
        body.lastName || "",
        body.phone || "",
        body.email || "",
        body.accountType || "",
        body.businessName || "",
        body.businessCategory || "",
        body.referralProgram || "",
        body.investorWanted || "No",
        body.verified ? "Yes" : "No",
      ];

    case "contact":
      return [
        ts,
        body.name || "",
        body.email || "",
        body.subject || "",
        body.message || "",
      ];

    case "volunteer":
      return [
        ts,
        body.name || "",
        body.email || "",
        body.phone || "",
        body.opportunity || "",
      ];

    case "investor-interest":
      return [
        ts,
        body.name || "",
        body.email || "",
        body.investmentRange || "",
        body.message || "",
        body.business || "",
      ];

    case "investor-register":
      return [
        ts,
        body.name || "",
        body.email || "",
        body.investmentRange || "",
        body.areasOfInterest || "",
      ];

    case "land-trust":
      return [
        ts,
        body.name || "",
        body.email || "",
        body.phone || "",
        body.interest || "",
        body.message || "",
      ];

    default:
      // Fallback: dump everything as JSON
      return [ts, JSON.stringify(body)];
  }
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Method Not Allowed" });
  }

  try {
    const body = req.body || {};
    const formType = (body.formType || "").toLowerCase();

    const tab = TAB_MAP[formType];
    if (!tab) {
      return res.status(400).json({ ok: false, error: `Unknown formType: "${formType}"` });
    }

    const auth = new google.auth.JWT(
      process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      undefined,
      (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
      ["https://www.googleapis.com/auth/spreadsheets"]
    );

    const sheets = google.sheets({ version: "v4", auth });

    const values = [buildRow(formType, body)];

    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.SHEET_ID,
      range: `${tab}!A:Z`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values },
    });

    return res.status(200).json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ ok: false, error: String(err) });
  }
}

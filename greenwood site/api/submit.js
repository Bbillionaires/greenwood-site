import { google } from "googleapis";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Method Not Allowed" });
  }

  try {
    const {
      name = "",
      phone = "",
      email = "",
      referralProgram = "",
      investorWanted = false,
      verified = false
    } = req.body || {};

    const auth = new google.auth.JWT(
      process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      undefined,
      (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
      ["https://www.googleapis.com/auth/spreadsheets"]
    );

    const sheets = google.sheets({ version: "v4", auth });

    const values = [[
      new Date().toISOString(),
      name,
      phone,
      email,
      referralProgram,
      investorWanted ? "Yes" : "No",
      verified ? "Yes" : "No"
    ]];

    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.SHEET_ID,
      range: `${process.env.SHEET_TAB || "Sheet1"}!A:Z`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values }
    });

    return res.status(200).json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ ok: false, error: String(err) });
  }
}

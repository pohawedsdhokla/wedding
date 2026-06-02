const SHEET_NAME = "RSVPs";
const RSVP_EMAIL = "pohawedsdhokla@gmail.com";

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Submitted At",
      "Name",
      "Email",
      "Attendance",
      "Guests",
      "Guest Names"
    ]);
  }

  const data = JSON.parse(e.postData.contents || "{}");

  sheet.appendRow([
    data.submittedAt || new Date().toISOString(),
    data.name || "",
    data.email || "",
    data.attendance || "",
    data.guests || "",
    data.guestNames || ""
  ]);

  MailApp.sendEmail({
    to: RSVP_EMAIL,
    subject: "New wedding RSVP from " + (data.name || "a guest"),
    body: [
      "A new RSVP was submitted.",
      "",
      "Name: " + (data.name || ""),
      "Email: " + (data.email || ""),
      "Attendance: " + (data.attendance || ""),
      "Guests: " + (data.guests || ""),
      "Guest Names: " + (data.guestNames || ""),
      "Submitted At: " + (data.submittedAt || new Date().toISOString())
    ].join("\n")
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

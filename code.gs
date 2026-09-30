/**
 * Google Apps Script for Nan Photography Portfolio
 * Handles submissions from both the "Contact" form and "Packages" modal.
 *
 * Setup Instructions:
 * 1. Open your Google Sheet > Extensions > Apps Script.
 * 2. Replace all code in Code.gs with this file.
 * 3. IMPORTANT: Select "testSendEmail" in the toolbar dropdown and click "Run".
 *    Google will prompt you to authorize permissions (click "Review Permissions" ->
 *    choose your Google account -> "Advanced" -> "Go to Portfolio (unsafe)" -> "Allow").
 *    Check your inbox to confirm the test email arrived!
 * 4. After authorizing, click Deploy > Manage deployments > Edit (pencil icon) >
 *    Version: New version > Deploy. (Always deploy a new version after editing).
 */

var RECIPIENT_EMAIL = "your-email@example.com"; // EDIT: Set your private notification email here in Google Apps Script editor

/**
 * Utility to strip CR/LF to prevent mail header injection.
 */
function stripNewlines(str) {
  return String(str || '').replace(/[\r\n]+/g, ' ').trim();
}

/**
 * Utility to HTML-escape user inputs if htmlBody is used.
 */
function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Utility to sanitize values before writing to Google Sheets to prevent formula injection.
 */
function safeCell(v) {
  var s = String(v == null ? '' : v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

/**
 * Run this function ONCE in the Apps Script editor to grant email sending permissions.
 */
function testSendEmail() {
  try {
    MailApp.sendEmail({
      to: RECIPIENT_EMAIL,
      subject: "Test Email - Nan Photography Portfolio",
      body: "If you are reading this, email permissions are authorized and working properly!"
    });
    console.log("Test email successfully sent via MailApp to " + RECIPIENT_EMAIL);
  } catch (err) {
    console.error("Test email failed: " + err.toString());
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Portfolio form endpoint is active." }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var params = (e && e.parameter) ? e.parameter : {};

    // 1. Honeypot check: silently ignore bots if botcheck field is filled
    if (params.botcheck && String(params.botcheck).trim().length > 0) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ignored" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Server-side throttling with CacheService
    var cache = CacheService.getScriptCache();
    var totalSubmissions = Number(cache.get("throttle_total") || 0);
    if (totalSubmissions >= 5) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "error", message: "Submission limit reached. Please wait a minute before trying again." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Input trimming, sanitization and length limits
    var rawEmail = String(params.email || "").trim().slice(0, 100);
    var cleanEmail = stripNewlines(rawEmail).toLowerCase();

    // Basic email validation
    var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailPattern.test(cleanEmail)) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "error", message: "Invalid email address format." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Per-email rate limit: 1 per email per 60 seconds
    var emailThrottleKey = "throttle_email_" + Utilities.base64EncodeWebSafe(cleanEmail);
    if (cache.get(emailThrottleKey)) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "error", message: "You recently submitted an enquiry. Please wait 60 seconds before submitting another." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var timestamp = new Date();
    var formType = stripNewlines(String(params.subject || "Website enquiry").slice(0, 80));
    var name = stripNewlines(String(params.name || "").slice(0, 100));
    var phone = stripNewlines(String(params.phone || "").slice(0, 30));
    var packageChosen = stripNewlines(String(params.package || "-").slice(0, 100));
    var message = String(params.message || "-").trim().slice(0, 2000);

    // Record throttling in cache
    cache.put("throttle_total", String(totalSubmissions + 1), 60);
    cache.put(emailThrottleKey, "1", 60);

    // 4. Append submission row to Google Sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Timestamp",
        "Form Type",
        "Full Name",
        "Phone Number",
        "Email",
        "Package Selected",
        "Message"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    sheet.appendRow([
      timestamp,
      safeCell(formType),
      safeCell(name),
      safeCell(phone),
      safeCell(cleanEmail),
      safeCell(packageChosen),
      safeCell(message)
    ]);

    // 5. Format notification email (stripped of newlines in subject and headers)
    var emailSubject = "Portfolio Enquiry: " + (name ? name : "New Lead") + " (" + formType + ")";
    var emailBody =
      "You received a new submission from your portfolio website:\n\n" +
      "Type: " + formType + "\n" +
      "Name: " + name + "\n" +
      "Phone: " + phone + "\n" +
      "Email: " + cleanEmail + "\n" +
      (packageChosen !== "-" ? "Package: " + packageChosen + "\n" : "") +
      (message !== "-" ? "Message:\n" + message + "\n\n" : "\n") +
      "Submitted at: " + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    var htmlBody =
      "<h3>New Portfolio Enquiry</h3>" +
      "<p><strong>Type:</strong> " + escapeHtml(formType) + "</p>" +
      "<p><strong>Name:</strong> " + escapeHtml(name) + "</p>" +
      "<p><strong>Phone:</strong> " + escapeHtml(phone) + "</p>" +
      "<p><strong>Email:</strong> " + escapeHtml(cleanEmail) + "</p>" +
      (packageChosen !== "-" ? "<p><strong>Package:</strong> " + escapeHtml(packageChosen) + "</p>" : "") +
      (message !== "-" ? "<p><strong>Message:</strong><br>" + escapeHtml(message).replace(/\n/g, '<br>') + "</p>" : "") +
      "<p><small>Submitted at: " + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") + "</small></p>";

    // 6. Send email notification securely
    try {
      var mailPayload = {
        to: RECIPIENT_EMAIL,
        subject: emailSubject,
        body: emailBody,
        htmlBody: htmlBody
      };
      if (cleanEmail && cleanEmail.indexOf("@") !== -1) {
        mailPayload.replyTo = cleanEmail;
      }
      MailApp.sendEmail(mailPayload);
    } catch (mailErr) {
      console.error("MailApp error: " + mailErr.toString());
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", message: "Enquiry received successfully" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("doPost error: " + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: "Unable to process request at this time." }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

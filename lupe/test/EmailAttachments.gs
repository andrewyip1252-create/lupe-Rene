// EmailAttachments.gs — kept as its own file, separate from Code.gs, so
// Code.gs itself stays a manageable size. These three files are embedded
// directly as base64 (confirmed necessary: the original Slack file links
// require being logged into the workspace to access at all, so
// GmailApp/UrlFetchApp running server-side could never fetch them
// directly — embedding avoids needing any external fetch entirely).
//
// getProspectsWelcomeAttachments() builds real Blob objects from this
// data for use in GmailApp.sendEmail's attachments option. Only used when
// the sender has explicitly opted in via the attachment toggles in Lupe's
// UI — never attached automatically.

var CONTACT_IMPORT_TEMPLATE_B64 = ""; // REDACTED — real xlsx template content lives only in the deployed Apps Script project, not in this public repo
var PROPERTIES_IMPORT_TEMPLATE_B64 = ""; // REDACTED — real xlsx template content lives only in the deployed Apps Script project, not in this public repo
var OUTLOOK_INTEGRATION_PDF_B64 = ""; // REDACTED — real PDF content lives only in the deployed Apps Script project, not in this public repo
var DEAL_IMPORT_CLOSED_TEMPLATE_B64 = ""; // REDACTED — real xlsx template content lives only in the deployed Apps Script project, not in this public repo
var DEAL_IMPORT_OPEN_TEMPLATE_B64 = ""; // REDACTED — real xlsx template content lives only in the deployed Apps Script project, not in this public repo

// keys: array of which attachments to include, e.g. ["contact","properties","outlook"]
function getProspectsWelcomeAttachments(keys) {
  var blobs = [];
  keys.forEach(function(key) {
    if (key === "contact") {
      blobs.push(Utilities.newBlob(Utilities.base64Decode(CONTACT_IMPORT_TEMPLATE_B64),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Contact Import Template.xlsx"));
    } else if (key === "properties") {
      blobs.push(Utilities.newBlob(Utilities.base64Decode(PROPERTIES_IMPORT_TEMPLATE_B64),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Properties Import Template.xlsx"));
    } else if (key === "outlook") {
      blobs.push(Utilities.newBlob(Utilities.base64Decode(OUTLOOK_INTEGRATION_PDF_B64),
        "application/pdf",
        "CRE OneSource - Outlook Integration IT Instructions.pdf"));
    } else if (key === "dealImportClosed") {
      blobs.push(Utilities.newBlob(Utilities.base64Decode(DEAL_IMPORT_CLOSED_TEMPLATE_B64),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Deal Import Template - Closed.xlsx"));
    } else if (key === "dealImportOpen") {
      blobs.push(Utilities.newBlob(Utilities.base64Decode(DEAL_IMPORT_OPEN_TEMPLATE_B64),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Deal Import Template - Open.xlsx"));
    }
  });
  return blobs;
}
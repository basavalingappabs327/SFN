/**
 * Sterilization Certificate document generator.
 *
 * Setup:
 * 1. Convert the uploaded DOCX template into a Google Docs document.
 * 2. Set TEMPLATE_DOC_ID below to the Google Docs document ID.
 * 3. In Apps Script: Project Settings → check "Show 'appsscript.json' manifest in editor".
 *    Copy appsscript.json from this folder (it includes script.external_request).
 * 4. In the editor, select authorizeExternalRequest and click Run. Accept the new
 *    "Connect to an external service" permission when Google asks.
 * 5. Deploy as a Web App (or Manage deployments → pencil → New version), Execute as: Me,
 *    and choose access allowed by your organisation.
 *
 * Word (.docx) export needs UrlFetchApp. If that permission is missing, PDF still
 * generates; run step 4 and create a new deployment version to enable Word.
 *
 * Formatting tip: In the Google Docs template, make only the placeholders you want bold
 * (e.g. {employeekannada}) Bold. replaceText keeps the formatting of each placeholder
 * range, so non-bold placeholders stay regular.
 */
const TEMPLATE_DOC_ID = "1ONeKHrxr6j-NXcViiAZxsohM_TLQLYFYmkTVe5iV2ec";

function doGet() {
  return jsonOutput_({ success: true, message: "Sterilization Certificate generator is running." });
}

function doPost(e) {
  try {
    if (!TEMPLATE_DOC_ID || TEMPLATE_DOC_ID.indexOf("PASTE_") === 0) {
      throw new Error("Set TEMPLATE_DOC_ID in Code.gs to your converted Google Docs template ID.");
    }

    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const operationDate = parseDate_(data.dateofoperation);
    const dob = parseDate_(data.dobofemployee);
    if (!operationDate) throw new Error("A valid date of operation is required.");
    if (!dob) throw new Error("A valid employee date of birth is required.");
    if (dob.getTime() > operationDate.getTime()) {
      throw new Error("Employee date of birth cannot be after the date of operation.");
    }

    // Backend-calculated fields; these are not entered in the HTML form.
    data.dateofoperation = formatDate_(operationDate);
    data.dobofemployee = formatDate_(dob);
    data.dateofsfn = formatDate_(new Date(operationDate.getFullYear(), operationDate.getMonth() + 1, 1));
    const ageYears = calculateAge_(dob, operationDate);
    // Employee age in Kannada, e.g. "46 ವರ್ಷಗಳು"
    data.employeeage = ageYears + " ವರ್ಷಗಳು";
    // Employee age in English, e.g. "46 years"
    data.employeeageeng = ageYears + " years";
    data.today = data.today || formatDate_(new Date());

    const template = DriveApp.getFileById(TEMPLATE_DOC_ID);
    const copyName = "Sterilization_Certificate_" +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") +
      "_" + Utilities.getUuid().slice(0, 8);
    const copy = template.makeCopy(copyName);
    const copyId = copy.getId();

    try {
      const doc = DocumentApp.openById(copyId);
      const body = doc.getBody();
      const replacements = buildReplacements_(data);

      // Insert real page breaks for {pagebreak} / {pagebreaK} before text replace.
      insertPageBreaks_(body);

      // Replace placeholders; formatting (bold etc.) on each placeholder in the
      // template is preserved for that range only — not forced on every Kannada field.
      replaceInContainer_(body, replacements);
      const header = doc.getHeader();
      if (header) replaceInContainer_(header, replacements);
      const footer = doc.getFooter();
      if (footer) replaceInContainer_(footer, replacements);

      doc.saveAndClose();

      const pdfBlob = DriveApp.getFileById(copyId).getAs(MimeType.PDF)
        .setName(copyName + ".pdf");

      // Google Docs files cannot be converted to DOCX with DriveApp.getAs().
      // Use the Google Drive API export endpoint instead. If UrlFetchApp is not
      // authorised yet, still return the PDF so the form is usable.
      let docxFileName = "";
      let docxBase64 = "";
      let docxError = "";
      try {
        const docxBlob = exportGoogleDocAsDocx_(copyId, copyName + ".docx");
        docxFileName = docxBlob.getName();
        docxBase64 = Utilities.base64Encode(docxBlob.getBytes());
      } catch (docxErr) {
        docxError = String(docxErr && docxErr.message ? docxErr.message : docxErr);
      }

      return jsonOutput_({
        success: true,
        pdfFileName: pdfBlob.getName(),
        pdfBase64: Utilities.base64Encode(pdfBlob.getBytes()),
        docxFileName: docxFileName,
        docxBase64: docxBase64,
        docxError: docxError,
        employeeage: data.employeeage,
        employeeageeng: data.employeeageeng,
        dateofsfn: data.dateofsfn
      });
    } finally {
      // Each request uses its own copy; trash it after conversion to avoid accumulating personal data.
      // The downloaded PDF and DOCX are returned as Base64, so the temporary Google Doc is not needed.
      try { DriveApp.getFileById(copyId).setTrashed(true); } catch (cleanupError) {}
    }
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err && err.message ? err.message : err) });
  }
}


/**
 * Run this once from the Apps Script editor (select this function, then Run).
 * That grants https://www.googleapis.com/auth/script.external_request so
 * Word export can call UrlFetchApp.fetch. Then publish a new Web App version.
 */
function authorizeExternalRequest() {
  UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/about?fields=kind", {
    method: "get",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
}

/** Export a Google Docs document to a real .docx file using Drive API v3. */
function exportGoogleDocAsDocx_(fileId, outputName) {
  const mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const url = "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(fileId) +
    "/export?mimeType=" + encodeURIComponent(mimeType);
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  if (status !== 200) {
    throw new Error("DOCX export failed (HTTP " + status + "): " + response.getContentText());
  }
  return response.getBlob().setContentType(mimeType).setName(outputName);
}

/**
 * Finds {pagebreak} / {pagebreaK} in the body and inserts a real PageBreak
 * at the end of the containing paragraph, then removes the placeholder text
 * during the normal replace step.
 * Page breaks cannot live inside tables/headers/footers.
 */
function insertPageBreaks_(body) {
  const patterns = ["\\{pagebreak\\}", "\\{pagebreaK\\}"];
  patterns.forEach(function(pattern) {
    var found = body.findText(pattern);
    while (found) {
      try {
        var parent = found.getElement().getParent();
        if (parent && parent.getType && parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
          parent.asParagraph().appendPageBreak();
        } else if (parent && parent.appendPageBreak) {
          parent.appendPageBreak();
        }
      } catch (e) {
        // Ignore if structure does not allow a page break (e.g. inside a table cell).
      }
      found = body.findText(pattern, found);
    }
  });
}

function buildReplacements_(d) {
  const result = {};
  const keys = [
    "doctor", "operation", "patient", "employee", "designation", "instiution",
    "institution", "institutionplace", "taluk", "district", "hospital",
    "dateofoperation", "dateofsfn", "employeeage", "employeeageeng", "patientage", "noofchildren",
    "1age", "2age", "employeekannada", "designationkannada",
    "institutionkannada", "institutionplacekannada", "patientkannada",
    "patientagekannada", "hospitalkannada", "noofchildrenkannada", "today",
    "dobofemployee", "talukkannada", "districtkannada",
    "basic", "scale", "patentkannada", "instiutionplacekannada",
    "instituitionplace", "pagebreak", "pagebreaK"
  ];
  keys.forEach(function(key) {
    result[key] = d[key] == null ? "" : String(d[key]);
  });

  // Match spelling variants that occur in the supplied template.
  result.institution = d.instiution || d.institution || "";
  result.instiution = d.instiution || d.institution || "";
  result.instituitionplace = d.institutionplace || "";
  result.instiutionplacekannada = d.institutionplacekannada || "";
  result.patentkannada = d.patientkannada || "";
  // pagebreak tokens are cleared after real PageBreak elements are inserted.
  result.pagebreak = "";
  result.pagebreaK = "";
  return result;
}

function replaceInContainer_(container, replacements) {
  Object.keys(replacements).forEach(function(key) {
    // Google Docs replaceText uses regex. Escape the literal braces in each placeholder.
    // Formatting already applied to the placeholder text in the template (e.g. Bold)
    // is kept on the replacement for that range only.
    const pattern = "\\{" + key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\}";
    container.replaceText(pattern, String(replacements[key] == null ? "" : replacements[key]));
  });
}

/** Accepts dd/mm/yyyy or yyyy-mm-dd and returns a local Date at midnight. */
function parseDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  const text = String(value).trim();
  let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) {
    const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    if (date.getFullYear() === Number(match[3]) &&
        date.getMonth() === Number(match[2]) - 1 &&
        date.getDate() === Number(match[1])) return date;
    return null;
  }
  match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    if (date.getFullYear() === Number(match[1]) &&
        date.getMonth() === Number(match[2]) - 1 &&
        date.getDate() === Number(match[3])) return date;
  }
  return null;
}

function calculateAge_(dob, operationDate) {
  let age = operationDate.getFullYear() - dob.getFullYear();
  const birthdayThisYear = new Date(operationDate.getFullYear(), dob.getMonth(), dob.getDate());
  if (operationDate < birthdayThisYear) age--;
  return age;
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function jsonOutput_(object) {
  return ContentService.createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}
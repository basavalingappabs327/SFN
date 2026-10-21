/*
  Configure this with your deployed Apps Script Web App URL ending in /exec.
*/
const APPS_SCRIPT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwmo1kb6-pQDYXQ0srxs3cEdAe_eX2y0FEKaxfRhpuuAVoQx2V83WVXrwRbwiZEuIVD/exec";

const form = document.getElementById("certificateForm");
const operationDateInput = document.getElementById("dateofoperation");
const dobInput = document.getElementById("dobofemployee");
const statusEl = document.getElementById("status");
const submitBtn = document.getElementById("generateBtn");
const downloadSection = document.getElementById("downloadSection");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const downloadDocxBtn = document.getElementById("downloadDocxBtn");

/** Holds the last successful generation result for on-demand downloads. */
let lastResult = null;

function formatDDMMYYYY(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${date.getFullYear()}`;
}
function localDateFromISO(value) {
  if (!value) return null;
  const parts = value.split("-").map(Number);
  if (parts.length !== 3 || parts.some(Number.isNaN)) return null;
  return new Date(parts[0], parts[1] - 1, parts[2]);
}
document.getElementById("today").value = formatDDMMYYYY(new Date());

// Employee age (Kannada/English) and dateofsfn are not displayed or calculated in the form.
// Code.gs calculates them from dateofoperation and dobofemployee on the server.

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!form.reportValidity()) return;

  if (!APPS_SCRIPT_WEB_APP_URL.startsWith("https://script.google.com/")) {
    statusEl.textContent = "Add your deployed Apps Script Web App /exec URL in script.js first.";
    return;
  }

  // Hide previous download buttons until a new successful generation.
  if (downloadSection) downloadSection.hidden = true;
  lastResult = null;

  const data = Object.fromEntries(new FormData(form).entries());
  // Relationship fields are intentionally not part of this form/template submission.
  delete data.relationship;
  delete data.relationshipkannada;
  const opDate = localDateFromISO(operationDateInput.value);
  const dob = localDateFromISO(dobInput.value);
  if (!opDate) {
    statusEl.textContent = "Please enter a valid date of operation.";
    return;
  }
  if (!dob) {
    statusEl.textContent = "Please enter a valid employee date of birth.";
    return;
  }
  if (dob > opDate) {
    statusEl.textContent = "Employee date of birth cannot be after the date of operation.";
    return;
  }

  data.dateofoperation = formatDDMMYYYY(opDate);
  data.dobofemployee = formatDDMMYYYY(dob);
  data.today = data.today || formatDDMMYYYY(new Date());
  // dateofsfn, employeeage, and employeeageeng are calculated by Code.gs, not supplied by the browser.
  // patientagekannada is submitted from the form as-is.

  submitBtn.disabled = true;
  statusEl.textContent = "Please wait — preparing your PDF and Word documents…";
  try {
    const response = await fetch(APPS_SCRIPT_WEB_APP_URL, {
      method: "POST",
      headers: {"Content-Type": "text/plain;charset=utf-8"},
      body: JSON.stringify(data)
    });
    if (!response.ok) throw new Error(`Server returned HTTP ${response.status}`);
    const result = await response.json();
    if (!result.success) throw new Error(result.error || "Document generation failed.");

    lastResult = result;
    if (downloadSection) downloadSection.hidden = false;
    if (downloadDocxBtn) {
      const hasDocx = Boolean(result.docxBase64);
      downloadDocxBtn.hidden = !hasDocx;
      downloadDocxBtn.disabled = !hasDocx;
    }
    if (result.docxBase64) {
      statusEl.textContent = "Documents are ready. Use the buttons below to download.";
    } else {
      statusEl.textContent = "PDF is ready. Word export needs Apps Script permission: run authorizeExternalRequest in the editor, accept access, then deploy a new Web App version.";
    }
  } catch (error) {
    console.error(error);
    statusEl.textContent = `Unable to generate documents: ${error.message}. Check the Web App deployment and template ID.`;
  } finally {
    submitBtn.disabled = false;
  }
});

if (downloadPdfBtn) {
  downloadPdfBtn.addEventListener("click", () => {
    if (!lastResult || !lastResult.pdfBase64) return;
    downloadBase64(
      lastResult.pdfBase64,
      lastResult.pdfFileName || "Sterilization_Certificate.pdf",
      "application/pdf"
    );
  });
}

if (downloadDocxBtn) {
  downloadDocxBtn.addEventListener("click", () => {
    if (!lastResult || !lastResult.docxBase64) return;
    downloadBase64(
      lastResult.docxBase64,
      lastResult.docxFileName || "Sterilization_Certificate.docx",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
  });
}

function downloadBase64(base64, filename, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  const blob = new Blob([bytes], {type: mimeType});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
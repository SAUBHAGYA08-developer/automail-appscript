/**
 * Returns email template by template name.
 * Reads from Templates sheet.
 */
function getTemplate(templateName) {
  // Utilizing our safe dynamic layout fetcher helper
  const sheet = getSheet(SHEETS.TEMPLATES);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const currentTemplate = values[i][0];

    // Added .toString().trim() validation protection to prevent whitespace matching issues
    if (currentTemplate && currentTemplate.toString().trim() === templateName) {
      return {
        subject: values[i][1] ? values[i][1].toString() : "",
        body: values[i][2] ? values[i][2].toString() : ""
      };
    }
  }

  throw new Error("Target template allocation failed. Key '" + templateName + "' was not found inside the Templates sheet ledger layout.");
}

/**
 * Replace placeholders in email text.
 * Safely accepts either a structured recruiter object OR a raw data sheet row array.
 */
function populateTemplate(text, recruiterData, signatureStr = "") {
  if (!text) return "";

  // 1. Dynamic Extraction Layer: Detect if input is a raw spreadsheet row array or parsed object
  let name = "";
  let company = "";
  let position = "";

  if (Array.isArray(recruiterData)) {
    name = recruiterData[COL.RECRUITER - 1] || "";
    company = recruiterData[COL.COMPANY - 1] || "";
    position = recruiterData[COL.POSITION - 1] || "";
  } else if (recruiterData && typeof recruiterData === "object") {
    name = recruiterData.name || recruiterData.recruiter || "";
    company = recruiterData.company || "";
    position = recruiterData.position || "";
  }

  // 2. Resolve Dynamic Greetings (Leveraging your project's native helper)
  const greeting = typeof getGreeting === "function" ? getGreeting(name) : name;

  // 3. Global Regex Replacement Processing Map
  return text
    .replace(/{{greeting}}/g, greeting)
    .replace(/{{name}}/g, name)
    .replace(/{{company}}/g, company)
    .replace(/{{position}}/g, position)
    .replace(/{{signature}}/g, signatureStr);
}
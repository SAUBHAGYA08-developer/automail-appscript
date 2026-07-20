function sendRecruiterEmail(row, rowNumber) {
  const recruiter = row[COL.RECRUITER - 1];
  const email = row[COL.EMAIL - 1];
  const company = row[COL.COMPANY - 1];
  const position = row[COL.POSITION - 1];
  const subjectOverride = row[COL.SUBJECT_OVERRIDE - 1];
  const status = row[COL.STATUS - 1];
  const threadId = row[COL.THREAD_ID - 1];

  const settings = getSettings();

  const templateName = selectTemplateName(status, position, company);
  const template = getTemplate(templateName);
  if (!template) {
    throw new Error("Target template '" + templateName + "' was not found in Templates sheet.");
  }

  const { subject, body } = buildEmailContent(
    template,
    { recruiter, company, position, subjectOverride },
    settings
  );

  const resumeBlob = getResumeAttachment(settings["Resume URL"]);

  // Check if existing thread exists
  if (threadId && threadId.toString().trim() !== "") {
    const cleanThreadId = threadId.toString().trim();
    const thread = GmailApp.getThreadById(cleanThreadId);

    if (thread) {
      // 📩 Purana thread -> Gmail API se Headers + Resume Attachment ke saath reply
      sendThreadReplyViaGmailApi(email, subject, body, thread, resumeBlob);
      return cleanThreadId;
    }
  }

  // ✉️ Bilkul naya email -> Fresh GmailApp.sendEmail
  GmailApp.sendEmail(email, subject, body, {
    attachments: [resumeBlob],
    name: settings["Sender Name"] || undefined
  });

  return findSentThreadId(email, subject);
}

function selectTemplateName(status, position, company) {
  if (status === STATUS.DETAILS_REQUESTED || status === "DETAILS_REQUESTED") {
    return "DETAILS_REQUESTED";
  }
  if (position && position.trim() !== "") {
    return "INITIAL_POSITION";
  }
  if (company && company.trim() !== "") {
    return "INITIAL_COMPANY";
  }
  return "INITIAL_DEFAULT";
}

function buildEmailContent(template, params, settings) {
  const { recruiter, company, position, subjectOverride } = params;

  // Signature resolution
  const activeUserEmail = Session.getActiveUser().getEmail();
  const userNameFallback = settings["Sender Name"] || activeUserEmail.split("@")[0];
  const signature = (settings["Signature"] && settings["Signature"].trim() !== "")
    ? settings["Signature"]
    : "Regards,\n\n" + userNameFallback;

  // Base Subject and Body
  let subject = (subjectOverride && subjectOverride.trim() !== "") 
    ? subjectOverride 
    : (template.subject || "");
  let body = template.body || "";

  // Dynamic values from Settings
  const totalExp = settings["Total Experience"] || "";
  const currentCtc = settings["Current CTC"] || "";
  const expectedCtc = settings["Expected CTC"] || "";
  const noticePeriod = settings["Notice Period"] || "";
  const location = settings["Current Location"] || "";
  const greeting = getGreeting(recruiter);

  // Replacements Map
  const replacements = [
    [/{{greeting}}/g, greeting],
    [/{{name}}/g, recruiter || ""],
    [/{{company}}/g, company || ""],
    [/{{position}}/g, position || ""],
    [/{{total_exp}}/g, totalExp],
    [/{{current_ctc}}/g, currentCtc],
    [/{{expected_ctc}}/g, expectedCtc],
    [/{{notice_period}}/g, noticePeriod],
    [/{{location}}/g, location],
    [/{{signature}}/g, signature]
  ];

  // Apply replacements to BOTH subject and body
  replacements.forEach(([regex, value]) => {
    subject = subject.replace(regex, value);
    body = body.replace(regex, value);
  });

  return { subject, body };
}

function getResumeAttachment(resumeUrl) {
  const fileId = getDriveFileId(resumeUrl);
  if (!fileId) {
    throw new Error("Missing or invalid 'Resume URL' configuration parameter value inside settings sheet.");
  }
  return DriveApp.getFileById(fileId).getBlob();
}

function findSentThreadId(email, subject) {
  Utilities.sleep(3000); // Allow Gmail search index to sync
  const threads = GmailApp.search('in:sent newer_than:1d', 0, 50);
  const cleanEmail = email.toString().trim().toLowerCase();

  for (const thread of threads) {
    const msg = thread.getMessages()[0];
    if (
      msg.getTo().toLowerCase().indexOf(cleanEmail) !== -1 &&
      msg.getSubject() === subject
    ) {
      Logger.log("FOUND CORRECT THREAD MATCH: " + thread.getId());
      return thread.getId();
    }
  }

  throw new Error("Unable to locate sent thread inside Gmail inbox search indexes.");
}

function sendFollowupEmail(threadId, attempt, row) {
  const settings = getSettings();

  const recruiter = row[COL.RECRUITER - 1];
  const email = row[COL.EMAIL - 1];
  const company = row[COL.COMPANY - 1];
  const position = row[COL.POSITION - 1];

  const templateKey = "FOLLOWUP" + attempt;
  const template = getTemplate(templateKey);

  if (!template) {
    Logger.log("Warning: Skipping follow-up execution loop because template '" + templateKey + "' was missing.");
    return;
  }

  // Shared Helper se content build hoga (Saare Placeholders + Signature supported)
  const { body } = buildEmailContent(
    template,
    { recruiter, company, position },
    settings
  );

  if (!threadId || threadId.toString().trim() === "") {
    throw new Error("Cannot send follow-up: Missing Thread ID for " + email);
  }

  const cleanThreadId = threadId.toString().trim();
  const thread = GmailApp.getThreadById(cleanThreadId);

  if (!thread) {
    throw new Error("Target thread ID not found in Gmail: " + cleanThreadId);
  }

  // Send via Gmail API
  sendThreadReplyViaGmailApi(email, null, body, thread, null);

  Logger.log("Follow-up #" + attempt + " successfully bumped in thread to: " + email);
}

function sendThreadReplyViaGmailApi(email, subject, body, thread, attachmentBlob) {
  const messages = thread.getMessages();
  if (!messages || messages.length === 0) {
    throw new Error("No messages found inside thread.");
  }

  const lastMessage = messages[messages.length - 1];
  const messageId = lastMessage.getHeader("Message-ID") || "";

  let replySubject = thread.getFirstMessageSubject();
  if (!replySubject.toLowerCase().startsWith("re:")) {
    replySubject = "Re: " + replySubject;
  }

  const boundary = "----=_Part_" + Date.now().toString(16);

  let mimeLines = [];

  // 1. Top-Level Headers
  mimeLines.push("To: " + email);
  mimeLines.push("Subject: " + replySubject);
  if (messageId) {
    mimeLines.push("In-Reply-To: " + messageId);
    mimeLines.push("References: " + messageId);
  }
  mimeLines.push("MIME-Version: 1.0");
  mimeLines.push('Content-Type: multipart/mixed; boundary="' + boundary + '"');
  mimeLines.push(""); // Mandatory empty line after top headers

  // 2. Text Body Part
  mimeLines.push("--" + boundary);
  mimeLines.push("Content-Type: text/plain; charset=UTF-8");
  mimeLines.push("Content-Transfer-Encoding: 8bit");
  mimeLines.push(""); // Mandatory empty line before body text
  mimeLines.push(body);
  mimeLines.push("");

  // 3. Attachment Part (Only if present)
  if (attachmentBlob) {
    const fileName = attachmentBlob.getName();
    const mimeType = attachmentBlob.getContentType() || "application/octet-stream";
    
    // Base64 encoding with RFC standard 76-character line wrapping
    const rawBase64 = Utilities.base64Encode(attachmentBlob.getBytes());
    const chunkedBase64 = rawBase64.match(/.{1,76}/g).join("\r\n");

    mimeLines.push("--" + boundary);
    mimeLines.push('Content-Type: ' + mimeType + '; name="' + fileName + '"');
    mimeLines.push('Content-Disposition: attachment; filename="' + fileName + '"');
    mimeLines.push("Content-Transfer-Encoding: base64");
    mimeLines.push(""); // Mandatory empty line before base64 content
    mimeLines.push(chunkedBase64);
    mimeLines.push("");
  }

  // 4. Closing Boundary
  mimeLines.push("--" + boundary + "--");

  const rawEmail = mimeLines.join("\r\n");

  // Send via Gmail API
const rawBytes = Utilities.newBlob(rawEmail, "text/plain; charset=UTF-8").getBytes();

  Gmail.Users.Messages.send({
    raw: Utilities.base64EncodeWebSafe(rawBytes),
    threadId: thread.getId()
  }, 'me');
}
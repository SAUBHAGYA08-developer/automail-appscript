function sendRecruiterEmail(row, rowNumber) {
  const recruiter = row[COL.RECRUITER - 1];
  const email = row[COL.EMAIL - 1];
  const company = row[COL.COMPANY - 1];
  const position = row[COL.POSITION - 1];
  const subjectOverride = row[COL.SUBJECT_OVERRIDE - 1];

  const settings = getSettings();

  // Unified fallback defaults for structural sheet integrity
  let templateName = "INITIAL_DEFAULT";
  if (position && position.trim() !== "") {
    templateName = "INITIAL_POSITION";
  } else if (company && company.trim() !== "") {
    templateName = "INITIAL_COMPANY";
  }

  const template = getTemplate(templateName);
  if (!template) {
    throw new Error("Target template '" + templateName + "' was not found in the Templates sheet layout.");
  }

  let subject = template.subject || "";
  let body = template.body || "";

  subject = subject
      .replace(/{{company}}/g, company || "")
      .replace(/{{position}}/g, position || "")
      .replace(/{{name}}/g, recruiter || "");

  if (subjectOverride && subjectOverride.trim() !== "") {
      subject = subjectOverride;
  }

  // Dynamic signature fallbacks based on the active user identity session
  const activeUserEmail = Session.getActiveUser().getEmail();
  const userNameFallback = settings["Sender Name"] || activeUserEmail.split("@")[0];
  
  const signature = settings["Signature"] && settings["Signature"].trim() !== ""
    ? settings["Signature"]
    : "Regards,\n\n" + userNameFallback;

  body = body
      .replace(/{{greeting}}/g, getGreeting(recruiter))
      .replace(/{{name}}/g, recruiter || "")
      .replace(/{{company}}/g, company || "")
      .replace(/{{position}}/g, position || "")
      .replace(/{{signature}}/g, signature);

  const fileId = getDriveFileId(settings["Resume URL"]);
  if (!fileId) {
    throw new Error("Missing or invalid 'Resume URL' configuration parameter value inside settings sheet.");
  }
  const resume = DriveApp.getFileById(fileId);

  GmailApp.sendEmail(email, subject, body, {
    attachments: [resume.getBlob()]
  });

  // Brief dynamic scale adjustment: process up to 50 items to ensure heavy loops don't drop target indices
  Utilities.sleep(3000);
  const threads = GmailApp.search('in:sent newer_than:1d', 0, 50);

  for (const thread of threads) {
    const msg = thread.getMessages()[0];

    if (
      msg.getTo().indexOf(email) !== -1 &&
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

  let body = template.body || "";

  // Dynamic signature resolution matching the initial email pipeline pattern
  const activeUserEmail = Session.getActiveUser().getEmail();
  const userNameFallback = settings["Sender Name"] || activeUserEmail.split("@")[0];
  const signature = settings["Signature"] && settings["Signature"].trim() !== ""
    ? settings["Signature"]
    : "Regards,\n\n" + userNameFallback;

  body = body
    .replace(/{{name}}/g, recruiter || "")
    .replace(/{{company}}/g, company || "")
    .replace(/{{position}}/g, position || "")
    .replace(/{{greeting}}/g, getGreeting(recruiter))
    .replace(/{{signature}}/g, signature);

  if (threadId && threadId.trim() !== "") {
    try {
      const cleanThreadId = threadId.trim();
      const thread = GmailApp.getThreadById(cleanThreadId);
      
      if (!thread) {
        throw new Error("Target thread ID element link not found: " + cleanThreadId);
      }

      const messages = thread.getMessages();
      if (!messages || messages.length === 0) {
        throw new Error("No usable historical elements exist in thread context.");
      }
      
      const lastMessage = messages[messages.length - 1];
      const messageId = lastMessage.getHeader("Message-ID");
      
      let originalSubject = thread.getFirstMessageSubject();
      if (!originalSubject.toLowerCase().startsWith("re:")) {
        originalSubject = "Re: " + originalSubject;
      }

      const rawEmail = [
        'To: ' + email,
        'Subject: ' + originalSubject,
        'In-Reply-To: ' + messageId,
        'References: ' + messageId,
        'Content-Type: text/plain; charset=UTF-8',
        '',
        body
      ].join('\r\n');

      // Threaded deployment execution utilizing standard platform integrations
      Gmail.Users.Messages.send({
        raw: Utilities.base64EncodeWebSafe(rawEmail),
        threadId: cleanThreadId
      }, 'me');
      
      Logger.log("Follow-up successfully bumped in thread to: " + email);
    } catch (e) {
      Logger.log("Thread error detail: " + e.toString());
      throw e; 
    }
  }
}
function testThread() {
  const thread = GmailApp.getThreadById("19f746f31d159726");

  Logger.log(thread.getFirstMessageSubject());

  const messages = thread.getMessages();

  Logger.log("Messages: " + messages.length);

  messages.forEach(m => {
    Logger.log("From: " + m.getFrom());
    Logger.log("To: " + m.getTo());
    Logger.log("----------------");
  });
}

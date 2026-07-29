function replytoquoterequest() {

  const threads = GmailApp.search(
  'is:unread newer_than:2d subject:"Quote Request"'
).concat(
  GmailApp.search(
    'is:unread newer_than:2d subject:"Demande de devis"'
  )
).slice(0, 10)//获取最近两天未读线程，从第0个线程开始，最多返回10个线程

 const templates =  SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName("Template");

  const email_tpls = templates.getDataRange().getValues();
  const tpls_en = email_tpls [4][3].replaceAll("{{name}}", name)
  const tpls_fr = email_tpls [3][3].replaceAll("{{name}}", name)

  if (threads.length === 0) {
    Logger.log("No unread emails.");
    return;
  }
  
  for (const thread of threads) {

  const message = thread.getMessages().pop();//获取最新message
  const subject = message.getSubject();
  Logger.log(subject)
  const sender = message.getFrom();
  const match = sender.match(/<(.+)>/);
  const email = match ? match[1] : sender;
 

  let body;
  if (subject.includes("Quote request")) {body = tpls_en}
  else if (subject.includes("Demande de devis")) {body = tpls_fr}
  Logger.log("after")
  Logger.log(body);

  if (body) {
   GmailApp.createDraft(
    email,
    subject,
    body
   );
  }
  //thread.markRead();

  }
}
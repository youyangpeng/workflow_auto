function replytoquoterequest() {
 
  const threads = GmailApp.search(
  'is:unread newer_than:2d subject:"Quote request"'
).concat(
  GmailApp.search(
    'is:unread newer_than:2d subject:"Demande de devis"'
  )
).slice(0, 10)//获取最近两天未读线程，从第0个线程开始，最多返回10个线程
 
 if (threads.length === 0) { Logger.log("No unread emails.");
    return;
  }

  const templates =  SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName("Template");
  const sheet = SpreadsheetApp.getActive().getSheetByName("Translators");
  const data = sheet.getDataRange().getValues();

  const email_tpls = templates.getDataRange().getValues();
  
  
  for (const thread of threads) {

    const message = thread.getMessages().pop();//获取最新message
    const subject = message.getSubject();
    const fromEmail = message.getFrom();
    const match = fromEmail.match(/<(.+)>/);
    const from = match ? match[1] : fromEmail;
    Logger.log(`from: ${from}`);
    let name = "";
    for (let i = 1; i < data.length; i++) {   // 从第2行开始，跳过标题
    const translator = data[i][0];  // A列
    const email = data[i][3];       // D列

      if (email === from) {
      name = translator.split("-")[0];
      break;   // 找到后停止循环
  
      }
    }
    const tpls_en = email_tpls [4][3].replaceAll("{{name}}", name)
    const tpls_fr = email_tpls [3][3].replaceAll("{{name}}", name)
    Logger.log(`name: ${name}`);
    let body;
    if (subject.includes("Quote request")) {body = tpls_en}
    else if (subject.includes("Demande de devis")) {body = tpls_fr}
    Logger.log(`body: ${body}`);
    message.createDraftReply(body)
 

    
  //thread.markRead();

  }
 }
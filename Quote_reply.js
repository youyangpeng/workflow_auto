function replyLatestEmail() {

  const threads = GmailApp.search('is:unread newer_than:2d(subject:"Quote Request" OR subject:"Demande de devis")',0,
10);//获取最近两天未读线程，从第0个线程开始，最多返回10个线程

  if (threads.length === 0) {
    Logger.log("No unread emails.");
    return;
  }

  for (const thread of threads) {

  const message = thread.getMessages().pop();//获取最新message

  const subject = message.getSubject();

  if (subject.includes("Quote Request")) {
  // 英文处理
  }
  else if (subject.includes("Demande de devis")) {
  // 法文处理
  }
  
  threads[0].markRead();

  }
}
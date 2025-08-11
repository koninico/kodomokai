function myFunction() {
  //紐づいているスプレッドシートの取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //シート「連絡先(添付あり)」の取得
  const sheet = ss.getSheetByName('連絡先（添付あり）'); //添付ファイルないときは（添付あり）を削除する！！
  const logSheetName = '送信ログ';

  // 各列の番号を定義
  const firstRow = 2;//連絡先データの開始行番号
  const nameCol = 1;  //名前の列番号
  const mailCol = 2;  //メールアドレスの列番号
  const subCol = 3;  //件名の列番号
  const bodyCol = 4;  //本文の列番号
  const repCol = 5; //担当者氏名の列番号
  const repaddCol = 6;  //担当者アドレスの列番号
  const requiredCol = 7;  //不要/必要の列番号

  //送信されるメールの数をカウントする変数
  let mailCount = 0;

  //連絡先数の取得（取得するアドレス数が何件か）
  const contactNum = sheet.getLastRow() - (firstRow - 1);

  //メールの件名の取得
  const subject = sheet.getRange(firstRow,subCol).getValue();
  Logger.log(subject);

  //添付するファイルのファイルID、添付ファイルがないときはコメントアウトする！！
  // const fileId = '';  
  // const attachmentFile = DriveApp.getFileById(fileId).getBlob();

  // ログシートの準備
  let logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
    logSheet.appendRow(["タイムスタンプ", "児童名", "メールアドレス", "結果", "担当者名", "担当者アドレス", "エラー内容"]);
  }

  //メール作成・メール送信
   for(let i = 0; i < contactNum; i++){
    //名前の取得
    const name = sheet.getRange(firstRow + i, nameCol).getValue();
    //送信メールアドレスの取得
    const to = sheet.getRange(firstRow + i,mailCol).getValue();

    //担当者の取得
    const rep = sheet.getRange(firstRow + i, repCol).getValue();

    //担当者のアドレス取得
    const repadd = sheet.getRange(firstRow + i, repaddCol).getValue();

    //不要・必要の取得
    const required = sheet.getRange(firstRow + i, requiredCol).getValue();
    
    //不要の場合はメール送信しない
    if(required === "不要"){
      Logger.log(name + "さんへのメール送信は不要です");
      continue;
    }
    
    //メール本文の作成
    const body = `${name}さんの保護者さま\n\n` + sheet.getRange(firstRow,bodyCol).getValue()+`\n※学年担当:${rep}\n${repadd}`;

    //メール返信先の設定
    const option = {
      replyTo: repadd, //返信先の設定
      // attachments: [attachmentFile], //添付ファイルの設定、添付ファイルなしの時はコメントアウトする！！
    };
    
    try {
      //メール送信
      GmailApp.sendEmail(to, subject, body, option);

      // ログに送信成功を出力
      Logger.log(name + "さん（" + to + "）へのメール送信が完了しました。");

      //メール送信件数をインクリメント
      mailCount++;

      // 成功ログ記録
      logSheet.appendRow([timestamp, name, to, "成功", rep, repadd, ""]);

    } catch (e) {
      // 送信エラー時のログ出力
      console.error(name + "さん（" + to + "）へのメール送信に失敗しました。エラー内容: " + e.message);
      // エラーログ記録
      logSheet.appendRow([timestamp, name, to, "エラー", rep, repadd, e.message]);
    }

  }
  //最終的なメール送信件数を出力
  Logger.log("送信されたメールの数" + mailCount);
}
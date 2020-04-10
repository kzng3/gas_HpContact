//シートを取得する。
var sheet = SpreadsheetApp.getActiveSheet();

function getMail() {

  // メール検索する文字列：今回は件名の【HP問合せ】
  var str = "HP問合せ";

  // 上記の文字列に合致するスレッドを取得(とりあえず100件取得します)
  var threads = GmailApp.search(str,0,100);

  // メッセージを取得する
  var messages = GmailApp.getMessagesForThreads(threads);

  for(var i = 0; i < messages.length; i++){
    for(var j = 0; j < messages[i].length; j++){

      //メッセージIDを取得(重複を防ぐため)
      var messageId = messages[i][j].getId();

      //もし、スプレッドシートに存在したら実行しない
      if(!hasId(messageId)){
       //メール受信日時を取得
       var mailDate = messages[i][j].getDate();

       // メッセージの本文をプレーンテキストで取得
       var body = messages[i][j].getPlainBody();

       //正規表現をつくる
       var regName = new RegExp('氏名 :' + '.*?' + '\r' );
       var regMail = new RegExp('メールアドレス :' + '.*?' + '\r' );
       var regTel = new RegExp('電話番号 :' + '.*?' + '\r' );
       var regMemo = new RegExp('備考 :' + '.*?' + '\r' );

       //正規表現をマッチさせたうえで、転記するときに、
       //セルに"お名前："などが入らないように、正規部分を削る
       var Name = body.match(regName)[0].replace("氏名：","");
       var Mail = body.match(regMail)[0].replace("メールアドレス :","");
       var Tel = body.match(regTel)[0].replace("電話番号 :","");
       var Memo = body.match(regMemo)[0].replace("備考 :","");

       //セルに行を追加する(ID、日時、氏名、アドレス、電話番号、備考)
       sheet.appendRow([messageID,mailDate,Name,Mail,Tel,Memo]);
       }
     }
  }      
}

// 同じIDのメールは転記しないようにするため、すでにIDがあるかどうか調べる関数
function hasId(id){  
   //今回は1列目にメールIDを入れていくので1列目から探す
    var data = sheet.getRange(1, 1,sheet.getLastRow(),1).getValues();
    var hasId = data.some(function(value,index,data){
   //コールバック関数
    return (value[0] === id);
  });
  return hasId;
}
/*
 *
 * @title: スプレッドシートのキーを返却します。
 *        -- グローバル変数の変わりにメソッドで取得
 * @return:キーを返却します。
 * 
 */
function getSpreadSheetKey(){
  
  return "スプレッドシートのキー";
}

/*
 *
 * @title: アクセストークンのUrlを返却します。
 * @return: アクセストークンを返却します。
 * 
 */
function getAccessToken(){

  return 'LINEボットのアクセストークン';
}

/*
 *
 * @title: LINEのメッセージ送信をトリガーに処理開始
 * @param: HTTPリクエスト
 * 
 */

//ポストで送られてくるので、送られてきたJSONをパース
function doPost(e) {
  
  //LINE Developersで取得したアクセストークンを入れる -- varでdoPostの外で宣言しても動かないのが原因＋Webアプリケーションとして公開しないと更新されない
  var CHANNEL_ACCESS_TOKEN = getAccessToken();
      
  var json = JSON.parse(e.postData.contents);
  
  //返信するためのトークン取得
  var reply_token= json.events[0].replyToken;
  
  if (typeof reply_token === 'undefined') {
    return;
  }

  //送られたメッセージ内容を取得
  var message = json.events[0].message.text;  
  
  // リンクと送られてきたらシートのリンクを返す
  if(message === "リンク"){
  
    postMessage(CHANNEL_ACCESS_TOKEN,reply_token,getSpreadSheetUrl());
    return;
  }
  
  //送られてきた項目名と金額をGoogleスプレッドシートに記載する
  var outputMessage =  writeSheet(message);
  
  // メッセージを組み立てる
  var outputMessage = makeMessage(outputMessage);
  
  //記載をすべて合算して返却する
  postMessage(CHANNEL_ACCESS_TOKEN,reply_token,outputMessage);


 }

/*
 *
 * @title: 項目をシートに書き込みます
 * @param: 書き込む内容
 * 
 */
function writeSheet(inputMessage) {
  
  // spreadsheetのキー
  var spreadSheetKey = getSpreadSheetKey();
  
  var insideFoodCell = "F1";
  var insideFoodColumn = "自炊費";
  var outsideFoodCell = "F2";
  var outsideFoodColumn = "外食費";
  var fixedCell = "F3";
  var fixedColumn = "固定費";
  var variableCell = "F4";
  var variableColumn = "変動費";
  var sumCell = "F5";
  var sumColumn = "合計";
  
  var per1ServeCell = "F7";
  var per1ServeColumn = "自炊費／一食";
  var perAll1ServeCell = "F8";
  var perAll1ServeColumn = "全食費／一食";
  var totalFoodCell = "F9";
  var totalFoodColumn = "食費合計";
  
  var formulaPartsStart = "=SUMIF(B:C,";
  var formulaPartsEnd = ",C:C)";

  var sumInsideFoodCell = "G1";
  var sumOutsideFoodCell = "G2";
  var sumFixedCell = "G3";
  var sumVariableCell = "G4";
  var sumAllCell = "G5";
  var sumTarget = "=SUM(" + sumInsideFoodCell + ":" + sumVariableCell + ")";
  var sumPer1ServeCell = "G7";
  var sumAllPer1ServeCell = "G8";
  var formulaPer1Serve = "=ROUND(G1/((3*(DATEDIF(TEXT(A2, \"yyyy/mm/01\") ,B2,\"D\")+1)-COUNTIF(B:B,\"外食費\"))*2),-1)";
  var formulaPerAll1Serve = "=ROUND(G9/((3*(DATEDIF(TEXT(A2, \"yyyy/mm/01\") ,B2,\"D\")+1))*2))";
  var sumTotalFoolCell = "G9";
  var formulaTotalFood = "=SUM(" + sumInsideFoodCell + "," + sumOutsideFoodCell +")";
  
  // targetのspreadSheetを取得
  var targetSpreadSheet = SpreadsheetApp.openById(spreadSheetKey);
  
  // 新しいシートを作る - シート名称が年月
  var sheetName = Utilities.formatDate(new Date(),"JST","yyyy/MM");
  // 過去データ入力のため、一時的にシート名称を変更
  //var sheetName = "2018/05";
  
  
  var targetSheet;
  if(!targetSpreadSheet.getSheetByName(sheetName)){
    
    targetSheet = targetSpreadSheet.insertSheet(sheetName);
    
    // 初回は式の埋め込みも行う。
    targetSheet.getRange(insideFoodCell).setValue(insideFoodColumn);
    targetSheet.getRange(outsideFoodCell).setValue(outsideFoodColumn);
    targetSheet.getRange(fixedCell).setValue(fixedColumn);
    targetSheet.getRange(variableCell).setValue(variableColumn);
    targetSheet.getRange(sumCell).setValue(sumColumn);
    targetSheet.getRange(per1ServeCell).setValue(per1ServeColumn);
    targetSheet.getRange(perAll1ServeCell).setValue(perAll1ServeColumn);
    targetSheet.getRange(totalFoodCell).setValue(totalFoodColumn);
    
    targetSheet.getRange(sumInsideFoodCell).setFormula(formulaPartsStart + insideFoodCell + formulaPartsEnd);
    targetSheet.getRange(sumOutsideFoodCell).setFormula(formulaPartsStart + outsideFoodCell + formulaPartsEnd);
    targetSheet.getRange(sumFixedCell).setFormula(formulaPartsStart + fixedCell + formulaPartsEnd);
    targetSheet.getRange(sumVariableCell).setFormula(formulaPartsStart + variableCell + formulaPartsEnd);
    targetSheet.getRange(sumAllCell).setFormula(sumTarget);
    targetSheet.getRange(sumPer1ServeCell).setFormula(formulaPer1Serve);
    targetSheet.getRange(sumAllPer1ServeCell).setFormula(formulaPerAll1Serve);
    targetSheet.getRange(sumTotalFoolCell).setFormula(formulaTotalFood);
    
  }else{
   
    targetSheet = targetSpreadSheet.getSheetByName(sheetName);
    
    // 食費計算用日付セット - 入力開始日
    targetSheet.getRange("A1").setValue("入力開始日");
    targetSheet.getRange("A2").setValue(Utilities.formatDate( new Date(), 'Asia/Tokyo', 'yyyy/MM/01'));
    // 食費計算用日付セット - 最新入力日
    targetSheet.getRange("B1").setValue("最新入力日");
    targetSheet.getRange("B2").setValue(Utilities.formatDate( new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));
    
    
  }  
    // 最終行を取得
  var lastRow = targetSheet.getLastRow() + 1;
  
  // LINEのメッセージを分割
  // 全角スペースがあった場合、半角スペースに置き換える
  inputMessage = inputMessage.replace(/　/g," ")

  // 半角スペース想定
  if(inputMessage.split(" ")){
      
    var items = inputMessage.split(" ");
  }
  
  var item = items[0];
  var price = items[1];
  var detail = "";
  
  // 備考がある場合のみ。 - outOfIndexException回避のため。
  if(items.length === 3){
  
    detail = items[2];
  }
  
  var inputYMDHM = Utilities.formatDate(new Date(),"JST","yyyy年MM月dd日 HH時mm分");
  var inputYMD = Utilities.formatDate(new Date(),"JST","yyyy年MM月dd日");
  var inputYM = Utilities.formatDate(new Date(),"JST","yyyy年MM日");
  var inputHM = Utilities.formatDate(new Date(),"JST","HH時mm分");
  
  // 1列目に日付 -- 仮に何も入力されてなかったら0.0になってしまうので注意
  targetSheet.getRange(lastRow,1).setValue(inputYMDHM);
  
  // 2列目に項目
  targetSheet.getRange(lastRow,2).setValue(item);
  
  // 3列目に額
  targetSheet.getRange(lastRow,3).setValue(price);
  
  // 4列目に備考
  targetSheet.getRange(lastRow,4).setValue(detail);
  
  // 合計額を取得する
  var sumInsideFood = 0;
  sumInsideFood = targetSheet.getRange(sumInsideFoodCell).getValue();
  
  var sumOutsideFood = 0;
  sumOutsideFood =   targetSheet.getRange(sumOutsideFoodCell).getValue();
  
  var sumFixed = 0;
  sumFixed =  targetSheet.getRange(sumFixedCell).getValue();
  
  var sumVariable = 0;
  sumVariable = targetSheet.getRange(sumVariableCell).getValue();
   
  var sumAll = 0;
  sumAll = targetSheet.getRange(sumAllCell).getValue();
  
  var sumPer1Serve = 0;
  sumPer1Serve = targetSheet.getRange(sumPer1ServeCell).getValue();
  
  var sumAllPer1Serve = 0;
  sumAllPer1Serve = targetSheet.getRange(sumAllPer1ServeCell).getValue();
  
  var sumTotalFool = 0;
  sumTotalFool = targetSheet.getRange(sumTotalFoolCell).getValue();
  
  // url取得
  var sheetUrl = targetSpreadSheet.getUrl();
  
  
  return {"inputYMD":inputYMD
          ,"inputHM":inputHM
          ,"inputYM":inputYM
          ,"item":item
          ,"price":price
          ,"sumInsideFood":sumInsideFood
          ,"sumOutsideFood":sumOutsideFood
          ,"sumFixed":sumFixed
          ,"sumVariable":sumVariable
          ,"sumAll":sumAll
          ,"sumPer1Serve":sumPer1Serve
          ,"sumAllPer1Serve":sumAllPer1Serve
          ,"sumTotalFool":sumTotalFool
          ,"sheetUrl":sheetUrl};
  
}

/*
 *
 * @title: スプレッドシートのurlを取得します。
 * @return: スプレッドシートのurl
 */
function getSpreadSheetUrl(){
  
  var spreadSheetKey = getSpreadSheetKey();
  
  // targetのspreadSheetを取得
  var targetSpreadSheet = SpreadsheetApp.openById(spreadSheetKey);
  
    // url取得
  var sheetUrl = targetSpreadSheet.getUrl();
  
  return sheetUrl;
}

/*
 *
 * @title: LINEに個別でメッセージを送る
 * @param: チャンネルアクセストークン（固定）
 * @param: リプライトークン（可変）
 * @param: メッセージの内容（可変）
 */
function postMessage(CHANNEL_ACCESS_TOKEN,reply_token,message){

  var line_endpoint = 'https://api.line.me/v2/bot/message/reply';
  
  try{
       // メッセージを返信    
       UrlFetchApp.fetch(line_endpoint, {
          'headers': {
              'Content-Type': 'application/json; charset=UTF-8',
              'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
           },
          'method': 'post',
          'payload': JSON.stringify({
          'replyToken': reply_token,
          'messages': [{
              'type': 'text',
              'text': message,
           }],
         }),
       });
       
    return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
  
  }catch(ex){
    
    console.log(ex);
  }
}

/*
 *
 * @title: LINEで全員にメッセージを送る
 * 
 * @param: チャンネルアクセストークン（固定）
 * @param: リプライトークン（可変）
 * @param: メッセージの内容（可変）
 */
function broadCastMessage(CHANNEL_ACCESS_TOKEN,reply_token,message){

  var line_endpoint = "https://api.line.me/v2/bot/message/push";
    
    
  try{
       // メッセージを返信    
       UrlFetchApp.fetch(line_endpoint, {
          'headers': {
              'Content-Type': 'application/json; charset=UTF-8',
              'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
           },
          'method': 'post',
          'payload': JSON.stringify({
          'replyToken': reply_token,
          'messages': [{
              'type': 'text',
              'text': message,
           }],
         }),
       });
       
    return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
  
  }catch(ex){
    
    console.log(ex);
  }
}

/*
 *
 * @title: LINEにを送るメッセージを組み立てる
 * @param: メッセージに必要な内容（可変）
 */
function makeMessage(outputMessage){  
  
  var br = "\n\r";
  var contents = "";
  var hr = "--------------------------------";
  
  contents += "【お知らせ】" + br + br;
  contents += "新たな出費を追加しました。" + br + br;
  contents += "年月日： " + outputMessage.inputYMD + br;
  contents += "時間  ： " + outputMessage.inputHM + br;
  contents += "項目  ：" + outputMessage.item + br;
  contents += "金額  ：" + insertComma(outputMessage.price) + "円" + br + br;
  contents += hr + br + br;
  contents += outputMessage.inputYM + "の合計額は以下の通りです。" + br + br;
  contents += "自炊費 ： "+insertComma(outputMessage.sumInsideFood) + "円"+ br;
  contents += "外食費 ： "+insertComma(outputMessage.sumOutsideFood) + "円"+ br;
  contents += "固定費 ： "+insertComma(outputMessage.sumFixed) + "円"+ br;
  contents += "変動費 ： "+insertComma(outputMessage.sumVariable) + "円"+ br;
  contents += "全合計 ：" +insertComma(outputMessage.sumAll) + "円"+ br + br;
  contents += hr + br + br;
  contents += "食費の部" + br + br;
  contents += "自炊費／一食： 約" + insertComma(outputMessage.sumPer1Serve) + "円" + br;
  contents += "全食費／一食：" + insertComma(outputMessage.sumAllPer1Serve) + "円" + br;
  contents += "食費合計 ：" + insertComma(outputMessage.sumTotalFool) + "円" + br +br;
  contents += "以上" + br;
  return contents;
  
  
} 

/*
 *
 * @title: 出力メッセージにカンマを入れる
 * @param: 変換前
 * @return: 変換後
 */
function insertComma(num) {
    return String(num).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,');
}

function urlShortener(url) {

  var API_KEY = 'MessagingApiのキー';

  var apiUrl  = 'https://www.googleapis.com/urlshortener/v1/url?key='+API_KEY;
  var payload = { longUrl: url };
  var options = {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
  var response = UrlFetchApp.fetch(apiUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('cannot shorten url.');
  } else {
    return JSON.parse(response).id;
  }
}
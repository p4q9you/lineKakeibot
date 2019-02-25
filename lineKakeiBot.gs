/*
  need to set properties - file > properties of project > properties of script
  need to Publish as web application
*/
 
/*
 * @title: write expense to sheet and reply line
 * @param: http request 
 */
function doPost(e) {
  
  // Insert the access token acquired with LINE Developers
  var CHANNEL_ACCESS_TOKEN = getAccessToken();
      
  var json = JSON.parse(e.postData.contents);
  
  // get replyToken Included in JSON
  var reply_token= json.events[0].replyToken;
  
  if (typeof reply_token === 'undefined') {
    return;
  }

  //get message Included in JSON
  var message = json.events[0].message.text;  
  
  // validate message
  var resultValidate = validateMessage(message);
  
  console.log(resultValidate);
  Logger.log(resultValidate);
  
  if(!resultValidate){
    postMessage(CHANNEL_ACCESS_TOKEN,reply_token,'正しい項目名を入力してください。');
    return;
  }
  
  // Return the link of the sheet as it is sent as a link
  if(message === "リンク"){
  
    postMessage(CHANNEL_ACCESS_TOKEN,reply_token,getSpreadSheetUrl());
    return;
  }
  
  // Write the item name and amount sent to Google Spreadsheet
  var outputMessage =  writeSheet(message);
  
  // Assemble the message
  var outputMessage = makeMessage(outputMessage);
  
  // Combine all the statements and return them
  postMessage(CHANNEL_ACCESS_TOKEN,reply_token,outputMessage);

 }

/*
 * @title: write sheet
 * @param: contents 
 */
function writeSheet(inputMessage) {
  
  // set valuefor building a sheet
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

  // get spreadsheet
  var targetSpreadSheet = getSpreadSheet();
  
  // Sheet name is year / month
  var sheetName = Utilities.formatDate(new Date(),"JST","yyyy/MM");
  
  // to debug
  //var sheetName = "2018/05";
  
  var targetSheet;
 
  if(!targetSpreadSheet.getSheetByName(sheetName)){
    
    // no sheet
    targetSheet = targetSpreadSheet.insertSheet(sheetName);
    
  }else{
   
    // is sheet
    targetSheet = targetSpreadSheet.getSheetByName(sheetName);

  }  

  // build sheet
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
    
  // meal cost - start
  targetSheet.getRange("A1").setValue("入力開始日");
  targetSheet.getRange("A2").setValue(Utilities.formatDate( new Date(), 'Asia/Tokyo', 'yyyy/MM/01'));
  // meal cost - end
  targetSheet.getRange("B1").setValue("最新入力日");
  targetSheet.getRange("B2").setValue(Utilities.formatDate( new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));
    
  // get last line
  var lastRow = targetSheet.getLastRow() + 1;
  
  // double-byte space　to a space
  inputMessage = inputMessage.replace(/　/g," ")

  // split message
  if(inputMessage.split(" ")){
      
    var items = inputMessage.split(" ");
  }
  
  // input message
  var item = items[0];
  var price = items[1];
  // optional
  var detail = "";
  if(items.length === 3){
  
    detail = items[2];
  }
  
  var inputYMDHM = Utilities.formatDate(new Date(),"JST","yyyy年MM月dd日 HH時mm分");
  var inputYMD = Utilities.formatDate(new Date(),"JST","yyyy年MM月dd日");
  var inputYM = Utilities.formatDate(new Date(),"JST","yyyy年MM日");
  var inputHM = Utilities.formatDate(new Date(),"JST","HH時mm分");
  
  // 1st column - date
  targetSheet.getRange(lastRow,1).setValue(inputYMDHM);
  
  // 2nd column - item
  targetSheet.getRange(lastRow,2).setValue(item);
  
  // 3rd column - price
  targetSheet.getRange(lastRow,3).setValue(price);
  
  // 4th column - detail
  targetSheet.getRange(lastRow,4).setValue(detail);
  
  // 5th column - to tally
  targetSheet.getRange(lastRow,8).setFormula('=SUBSTITUTE(SUBSTITUTE(LEFT(A' + lastRow + ',10),"年","/"),"月","/")');
  
  // get total amount
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
  
  // get url
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
 * @title: Send messages individually with LINE
 * @param: access token
 * @param: reply token
 * @param: content
 */
function postMessage(CHANNEL_ACCESS_TOKEN,reply_token,message){

  var line_endpoint = 'https://api.line.me/v2/bot/message/reply';
  
  try{
       // send message    
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
 * @title: Send a message to everyone on LINE
 * @param: access token
 * @param: reply token
 * @param: content
 */
function broadCastMessage(CHANNEL_ACCESS_TOKEN,reply_token,message){

  var line_endpoint = "https://api.line.me/v2/bot/message/push";
     
  try{
       // send message    
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
 * @title: Assemble message to send to LINE
 * @param: content
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
 * @title: Put a comma in the output message
 * @param: before
 * @return: after
 */
function insertComma(num) {
    return String(num).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,');
}

/*
 * Daily expense notification
 */
function pushMessage() {

     getSheetOfSpreadSheet().getRange("B2").setValue(Utilities.formatDate( new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));
     
     var br = "\n\r";
     var message = getSheetOfSpreadSheet().getRange('C2').getValue();
     var totalVal = getSheetOfSpreadSheet().getRange('G5').getValue();
     var remainDayCount = getSheetOfSpreadSheet().getRange('D2').getValue();
  
       // var line_endpoint = 'https://api.line.me/v2/bot/message/push';
       var line_endpoint = 'https://api.line.me/v2/bot/message/multicast';

       // send message    
       UrlFetchApp.fetch(line_endpoint, {
          'headers': {
              'Content-Type': 'application/json; charset=UTF-8',
              'Authorization': 'Bearer ' + getAccessToken()
           },
         'method': 'post',
          'payload': JSON.stringify({
            'to':[getHusbandUserId(),getWifeUserId()],
            //'to':getWifeUserId(),
            'messages': [{
              'type': 'text',
              'text': '本日の出費 : ' + insertComma(message) + '円' + br 
                    + '今月の合計出費:' + insertComma(totalVal) + '円' + br
                    + '今月は残り' + remainDayCount + '日です。',
            }],
         }),
       });   
}

/*
 * @title: validate message
 */
function validateMessage(target){

  var items;
  
  // double-byte space　to a space
  target = target.replace(/　/g," ")

  // split message
  if(target.split(" ")){
      
    items = target.split(" ");
  }
  
  // input message
  var item = items[0];
  
  // error flag
  var validateErrorFlg = true;
  
  if(item != "リンク" 
     && item != "自炊費"
     && item != "外食費"
     && item != "固定費"
     && item != "変動費"){
    
    validateErrorFlg = false;
  }
  
  return validateErrorFlg;
  
}

function getSheetOfSpreadSheet(){

 // Sheet name is year / month
  var sheetName = Utilities.formatDate(new Date(),"JST","yyyy/MM");
  
  return getSpreadSheet().getSheetByName(sheetName);
}

/*
 * @title: get url of screadSheet
 */
function getSpreadSheetUrl(){
  
  return SpreadsheetApp.openById(getSpreadSheetKey()).getUrl();
}

/*
 * @title: return key of Spreadsheet 
 */
function getSpreadSheetKey(){
  
  return PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_KEY'); 
}

/*
 * @title:return Spreadsheet 
 */ 
function getSpreadSheet(){
 return SpreadsheetApp.openById(getSpreadSheetKey()); 
}

/*
 * @title:return userid of husband 
 */ 
function getHusbandUserId(){
 return PropertiesService.getScriptProperties().getProperty('HUSBAND_USER_ID');  
} 

/*
 * @title:return userid of wife 
 */ 
function getWifeUserId(){
 return PropertiesService.getScriptProperties().getProperty('WIFE_USER_ID');  
}


/*
 * @title: return access token
 */
function getAccessToken(){

  return PropertiesService.getScriptProperties().getProperty('ACCESS_TOKEN');  
}
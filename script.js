//by Terence Siu
/**
1. can process int
2. price should be number
3. total price show in email
 */
function onOpen() {
  var menu = [{name: 'gen form', functionName: 'setUpConference_'}];
  SpreadsheetApp.getActive().addMenu('Generate', menu);
}
function setUpConference_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('products');
  var range = sheet.getDataRange();
  var values = range.getValues();
  setUpForm_(ss, values);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
}
function setUpForm_(ss, values) {
  var schedule = {};
  var form = FormApp.create('山城士多 訂購表 2017');
  form.setDescription('我地係一群中大學生...遲D再加');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.addTextItem().setTitle('姓名 Name').setRequired(true);
  form.addTextItem().setTitle('聯絡電話 Contact number').setRequired(true);
  form.addTextItem().setTitle('電郵 Email address').setRequired(true);
  form.addTextItem().setTitle('書院 College').setRequired(true);
  form.addMultipleChoiceItem().setTitle('你.... You are...').setChoiceValues(['是住在中大的 living in CUHK campus','不是住在中大的 not living in CUHK campus']).setRequired(true);
  form.addTextItem().setTitle('其他事項/意見 Other suggestions/advice').setRequired(false);
  form.addTextItem().setTitle('希望山城士多增加什麼類型的貨品? Want us to add more items?').setRequired(false);
  form.addMultipleChoiceItem().setTitle('希望日後收到山城士多的最新資訊嗎? Wish to receive our latest news?XD')
  .setChoiceValues(['好啊 Yes ar!','不了 No la~']).setRequired(true);
  
  var item = ss.getSheetByName("products");
  var brand = item.getRange(2, 2, item.getMaxRows() - 1).getValues();
  var ebrand = item.getRange(2, 3, item.getMaxRows() - 1).getValues();
  var itemValues = item.getRange(2, 4, item.getMaxRows() - 1).getValues();
  var eitemValues = item.getRange(2, 5, item.getMaxRows() - 1).getValues();
  var price = item.getRange(2, 7, item.getMaxRows() - 1).getValues();
  var quantity = item.getRange(2, 8, item.getMaxRows() - 1).getValues();
  var quotaValues = item.getRange(2, 10, item.getMaxRows() - 1).getValues();
  var categories = item.getRange(2, 13, item.getMaxRows() - 1).getValues();
  
  var precat='';
  for(var i = 0; i < itemValues.length; i++)
    if (quotaValues[i][0] != "")
    {
      var tmp = [];
      var max = Number(quotaValues[i][0])>5?5:Number(quotaValues[i][0]);
      for(var j = 0; j <=max; j++) {
        tmp[j]=j;}
      
      var cat = categories[i][0];
      if (precat!=cat) {form.addSectionHeaderItem().setTitle(cat);precat=cat;}
      var mc=form.addListItem();
      //var mc = form.getItems(FormApp.ItemType.LIST)[0].asListItem();
      mc.setTitle("["+brand[i][0]+ebrand[i][0] +"] "+itemValues[i][0]+" "+ eitemValues[i][0]+" $"+Number(price[i][0])+" / "+quantity[i][0]);
      mc.setChoiceValues(tmp);
    }    
}
function onFormSubmit(e) {
  var user = {name: e.namedValues['姓名 Name'][0], email: e.namedValues['電郵 Email address'][0]};
  var response = [];
  var totalprice = 0;
  var pricess = [];
  var values = SpreadsheetApp.getActive().getSheetByName('products')
     .getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var brand = session[1];
    var ebrand = session[2];
    var name = session[3];
    var ename = session[4];
    var price = session[6];   
    var quantity = session[7];
    timeslot = "["+brand+ebrand+"] "+name+" "+ename +" $"+price+" / "+quantity;
      Logger.log(timeslot);
    if (e.namedValues[timeslot] && Number(e.namedValues[timeslot]) >0) {
      response.push(session);
      pricess.push(parseInt(e.namedValues[timeslot][0],10));
      totalprice+=parseInt(e.namedValues[timeslot][0],10)*price;
      //Logger.log(parseInt(e.namedValues[timeslot][0],10));
      //Logger.log(price);
    }

  }
  sendDoc_(user, response,pricess,totalprice);
}

function sendDoc_(user, response,price,totalprice) {
  var content='';
  for (var i = 0; i < response.length; i++) {
     content=content+('['+response[i][1]+' '+response[i][3]+' '+response[i][4]+' $'+response[i][5]+'/'+response[i][7])+'] x'+price[i]+'\n';
  }
  MailApp.sendEmail({
    to: user.email,
    subject: '【山城士多】 訂貨確認',
    body: ' 親愛的'+user.name+': \n\n我們已經收到你在山城士多的訂單，謝謝支持。請核對以下的訂購項目及款項：\n\n'+content+'\n\n總價錢： $'+totalprice+'\n如無問題的話，請按照以下步驟付款：\n------------\n方法一 銀行過數：\n【付費】把訂購的總數款項過到以下戶口 :\n290-872795-668 恒生(Chan Ho Ching, Cheung Lung)\n【入數紙】在25/1/2018 23:59 (特別注明之產品除外)把入數紙send email到 cuhk.buy@gmail.com\n\n方法二 山城角樂付款：\n【付款】在25/1/2018 20:30前到山城角樂付現金，保存收據（不設找續！不設找續！）\n\n -------------\n【取貨】在29/1/2018 12:30-18:30  在 山城角樂(素食餐廳旁)\n\n如有任何問題，歡迎回覆此 E-mail 或 Facebook Inbox 我們！\n再次感謝你對山城士多的支持，更重要的是謝謝你願意認識和支持香港小店！\n\n祝生活愉快！\n山城士多上', });
}
function updateForm(){
  var form = FormApp.openByUrl("https://docs.google.com/forms/d/1LaI65ItguPLN4GCIhdPRrNxznII2yk5iQLFLIZ1IA_Y/edit");
  var ss = SpreadsheetApp.getActive();
  var item = ss.getSheetByName("products");
  var brand = item.getRange(2, 2, item.getMaxRows() - 1).getValues();
  var name = item.getRange(2, 4, item.getMaxRows() - 1).getValues();
  var ename = item.getRange(2, 5, item.getMaxRows() - 1).getValues();
  var price = item.getRange(2, 7, item.getMaxRows() - 1).getValues();
  var quantity = item.getRange(2, 8, item.getMaxRows() - 1).getValues();
  var quotaValues = item.getRange(2, 10, item.getMaxRows() - 1).getValues();
  var categories = item.getRange(2, 13, item.getMaxRows() - 1).getValues();
  
  for(var i = 0; i < name.length; i++)
    
    if(quotaValues[i][0] != "")
    {
      var tmp = [];
      var max = Number(quotaValues[i][0])>5?5:Number(quotaValues[i][0]);
      for(var j = 0; j <=max; j++) {
        tmp[j]=j;}
      //var mc=form.addListItem()
      var mc = form.getItems(FormApp.ItemType.LIST)[i].asListItem();
      mc.setTitle("["+brand[i][0]+"] "+name[i][0]+" "+ename[i][0]+" $"+price[i][0]+" / "+quantity[i][0]);
      mc.setChoiceValues(tmp);
    }    
}
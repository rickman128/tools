// ---------------------------------------------
// 入力値の取得関数
// ---------------------------------------------
function doGet(e){
  return HtmlService.createHtmlOutputFromFile('index');
}

// ---------------------------------------------
// 時間指定イベントチェック関数
// ---------------------------------------------
function checkData(frm){

  /* 入力パラメータ */
  // 業者名
  var user1 = frm['user'];
  // メールアドレス
  var mail = frm['mail'];
  // 日付
  var date1 = new Date(frm['date']); 
  // 開始時間
  var timeFrom1 = frm['timeFrom'];
  // 終了時間
  var timeTo1 = frm['timeTo'];
  // 用件
  var business1 = frm['business'];
  // 機器名/システム
  var system1 = frm['system'];

  /* 入力値のチェック */
  if (!checkInputData(user1, mail, date1, timeFrom1, timeTo1, business1, system1)){  
    // 入力エラーあり
    // ざっくりすぎ。メッセージをわかるように戻り値を設定
    return "入力値が正しくありません。";
  }
  else{
    var title = "【" + user1 + "】" + business1 + "：" + system1;
    Logger.log("------------------------------------------------------");
    Logger.log("   ■入力値   ");
    Logger.log("------------------------------------------------------");
    Logger.log(title);
    Logger.log(date1.getYear() + "年" + (date1.getMonth() + 1) + "月" + date1.getDate() + "日" + " " + timeFrom1 + "～" + timeTo1);
    Logger.log("------------------------------------------------------");
    
    // 同日の開始時間、終了時間Dateオブジェクト作成
    var time1 = new Date(date1.getYear(), date1.getMonth(), date1.getDate(), timeFrom1.substring(0,timeFrom1.indexOf(":")), timeFrom1.substring(timeFrom1.indexOf(":") + 1));
    var time2 = new Date(date1.getYear(), date1.getMonth(), date1.getDate(), timeTo1.substring(0,timeTo1.indexOf(":")), timeTo1.substring(timeTo1.indexOf(":") + 1));
     
    //Logger.log(time1.getYear() + "年" + (time1.getMonth() + 1) + "月" + time1.getDate() + "日");
    //Logger.log(time1.getHours() + "時" + time1.getMinutes() + '分　～　' + time2.getHours() + "時" + time2.getMinutes() + "分");
        
    var cal = CalendarApp.getCalendarById('sr5fhcn2bvunjrto8gmdt9cr28@group.calendar.google.com');
    var result = '';
    
    /* 登録する前に重複するイベントがないかチェック */
    // 範囲内のイベント取得
    var evts = cal.getEvents(time1, time2);
     
    // 重複するイベントあり
    if (evts.length > 0){
      Logger.log("------------------------------------------------------");
      Logger.log("   ■重複するイベント");
      Logger.log("------------------------------------------------------");
        for (var i = 0; i < evts.length; i++){
        Logger.log(evts[i].getTitle());
      }
      Logger.log("------------------------------------------------------");
        result ='申し訳ありません。その時間は既に予約済みです。';
    // 重複するイベントなし
    } else {
      Logger.log("createEvent前");
      // イベント作成
      result = createEvent(title, time1, time2, user1, mail, cal);    
      Logger.log("createEvent後");    
      
      // メール送信
      GmailApp.sendEmail('kendo2413@gmail.com', 'ME予約登録メール', title)
    }
    
    return result;
  }
}

// ---------------------------------------------
// 入力チェック
// ---------------------------------------------
function checkInputData(user, mail, date, timeFrom, timeTo, business, system){
  // 未入力？
  if ((user == "") ||
    (mail == "") ||
    (date == "") ||
    (timeFrom == "") ||
    (timeTo == "") ||
    (business == "") ||
    (system == "")){
      return false;
  }
  // 日付
  // この時点でnew Dateしているためチェックするには遅い
  // 開始・終了時間
  else if ((!chkTime(timeFrom)) || (!chkTime(timeTo))){
    return false;
  }
  // メールアドレス
  else if (!chkMail(mail)){
    return false;
  }
  else{
    return true;
  }   
}

/****************************************************************
* 機　能： 入力された値が時間でHH:MM形式になっているか調べる
* 引　数： str　入力された値
* 戻り値： 正：true　不正：false
****************************************************************/
function chkTime(str) {
    // 正規表現による書式チェック
    if(!str.match(/^\d{2}\:\d{2}$/)){
        return false;
    }
    var vHour = str.substr(0, 2) - 0;
    var vMinutes = str.substr(3, 2) - 0;
    if(vHour >= 0 && vHour <= 24 && vMinutes >= 0 && vMinutes <= 59){
        return true;
    }else{
    }
}

/****************************************************************
* 機　能： 入力された値がメールアドレスか調べる
* 引　数： str　入力された値
* 戻り値： 正：true　不正：false
****************************************************************/
function chkMail(str) {
   ml = /.+@.+\..+/; // チェック方式

   if(!str.match(ml)) {
     return false;
   }
   return true;
}

// ---------------------------------------------
// イベント作成関数
// ---------------------------------------------
function createEvent(title, timeFrom, timeTo, user, mail, cal){
    // カレンダーに新規予定登録
    cal.createEvent(title, timeFrom, timeTo,
                          {description:user + '様連絡先：' + mail, guests:mail});
    var strMinFrom = timeFrom.getMinutes();
    var strMinTo = timeTo.getMinutes();
    if (strMinFrom == 0){
      strMinFrom = '00';
    }
    if (strMinTo == 0){
      strMinTo = '00';
    }
    result = timeFrom.getYear() + "年" + (timeFrom.getMonth() + 1) + 
      "月" + timeFrom.getDate() + "日" + timeFrom.getHours() + ":" + strMinFrom + " ～ " + 
        timeTo.getHours() + ":" + strMinTo + 'に予約しました。（ブラウザーをリロードすると、カレンダーが更新されます）';
        
  return result;
}

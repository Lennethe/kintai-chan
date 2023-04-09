function getUserName(token, userId){
  const options = {
    "method" : "get",
    "contentType": "application/x-www-form-urlencoded",
    "payload" : { 
      "token": token
    }
  };
  
  const url = "https://slack.com/api/users.list";
  const response = UrlFetchApp.fetch(url, options);
  const members = JSON.parse(response).members;
    
  for (const member of members) {    
    //削除済、botユーザー、Slackbotを除く
    if (!member.deleted && !member.is_bot && member.id === userId) {
      let id = member.id;
      let real_name = member.real_name; //氏名(※表示名ではない)
      return real_name;
    }
    
  }
  
}

function findRow(sheet, val, col){
  const dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  for(var i=1;i<dat.length;i++){
    if(dat[i][col] === val){
      return i+1;
    }
  }
  return 0;
}
function randomGet(arr) {
  var index = Math.floor(Math.random() * arr.length)
  return arr[index];
}
function sumPunchTimeAndSecond(punchTime, sumtime){
  return ((new Date() - new Date(punchTime))) + parseInt(sumtime)
}


function doPost(e) {
  const token = PropertiesService.getScriptProperties().getProperty("token")
  const slackApp = SlackApp.create(token);
  const userNameCol = 1
  const idCol = 2
  const stateCol = 3
  const workTimeCol = 4
  const awayTimeCol = 5
  const punchTimeCol = 6
  const workStartedAtCol = 7


  // const command = "/test"
  // const userId = "U02H1GDFBM2"
  // const userName = "yoshida"
  // const channelId = "C03JC0JLKHC";
  // const text = "退勤"
  const param = e.parameter;
  const command = param.command
  const userId = param.user_id
  const userName = getUserName(token, userId);
  const channelId = param.channel_id
  const text = param.text
  const datetime = new Date()


  const kinmuJoutaiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('勤務状態');
  const kintaiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('勤怠');
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('勤怠ログ');


  const templateArr = [["No", "ユーザー名", "ユーザーID", "状態", "労働時間", "離席時間", "打刻日時", "出勤日時"]]
  const kintaiTemplateArr = [["No", "ユーザー名", "ユーザーID", "労働時間", "離席時間", "出勤日時", "退勤日時"]]
  const logArr = [[logSheet.getLastRow(), userName, userId, text, datetime]]



  // 登録していなければ状態を勤務外で追加。userRowNumの取得。
  let userRowNum = findRow(kinmuJoutaiSheet, userId, idCol)
  if (userRowNum === 0){
    userRowNum = kinmuJoutaiSheet.getLastRow() + 1
    const assignedArr = [[userRowNum-1, userName, userId, "勤務外", 0, 0, datetime, ""]]
    kinmuJoutaiSheet.getRange(userRowNum, 1, 1, assignedArr[0].length).setValues(assignedArr)
  }
  const userRange = kinmuJoutaiSheet.getRange(userRowNum, 1, templateArr.length, templateArr[0].length)
  const userValues = userRange.getValues()
  const userState = userValues[0][stateCol];

  const kintaiRange = kintaiSheet.getRange(kintaiSheet.getLastRow() + 1, 1, kintaiTemplateArr.length, kintaiTemplateArr[0].length)
  const logRange = logSheet.getRange(logSheet.getLastRow() + 1, 1, logArr.length, logArr[0].length)

  // ログを取る
  logRange.setValues(logArr);


  function setShukkin(){
    userValues[0][stateCol] = "勤務中"
    userValues[0][workTimeCol] = 0
    userValues[0][awayTimeCol] = 0
    userValues[0][punchTimeCol] = datetime
    userValues[0][workStartedAtCol] = datetime
    userRange.setValues(userValues)
  }
  function setWorkTime(state){
    userValues[0][stateCol] = state
    userValues[0][workTimeCol] = sumPunchTimeAndSecond(userValues[0][punchTimeCol], userValues[0][workTimeCol]) 
    userValues[0][punchTimeCol] = datetime
    userRange.setValues(userValues)
  }
  function setAwayTime(state){
    userValues[0][stateCol] = state
    userValues[0][awayTimeCol] = sumPunchTimeAndSecond(userValues[0][punchTimeCol], userValues[0][awayTimeCol]) 
    userValues[0][punchTimeCol] = datetime
    userRange.setValues(userValues)
  }
  function setTaikin(){
    // スプレッドシートの日時秒変換を用いるため、一旦記入してから結果を読み込む
    const kintaiWriteArr = [[
      kintaiSheet.getLastRow(),
      userName,
      userId,
      userValues[0][workTimeCol]/(24*60*60*1000),
      userValues[0][awayTimeCol]/(24*60*60*1000),
      userValues[0][workStartedAtCol],
      datetime
    ]]
    kintaiRange.setValues(kintaiWriteArr)
  }

  function postMessage(state){
    if (state === "出勤"){
      slackApp.postMessage(channelId, datetime.getHours() + "時" + datetime.getMinutes() + "分に" + userName + "さんが出勤されました。")
    }
    else if (state === "休憩"){
      slackApp.postMessage(channelId, datetime.getHours() + "時" + datetime.getMinutes() + "分に" + userName + "さんが休憩に入りました。")
    }
    else if (state === "再開"){
      slackApp.postMessage(channelId, datetime.getHours() + "時" + datetime.getMinutes() + "分に" + userName + "さんが勤務を再開しました。")
    }
    else if (state === "退勤"){
      const kintaiReadArr = kintaiRange.getValues()
      const workTime = kintaiReadArr[0][3]
      const awayTime = kintaiReadArr[0][4]
      slackApp.postMessage(channelId, 
        userName + "さんの今日の勤怠時間です。\n" + 
        "今日働いた時間は" + workTime.getHours() + "時間" + workTime.getMinutes() + "分で、\n" + 
        "休憩してた時間は" + awayTime.getHours() + "時間" + awayTime.getMinutes() + "分です。"
      )
    }
  }
  function setResponse(state){
    let ans = [];
    if (state === "出勤"){
      ans = [
        "おはようございます！今日も素敵な一日にしましょう！",
        "おはよう！今日も最高の笑顔で頑張ろう！",
        "朝だー！元気にいきましょう！",
        "おはようございまーす！今日も一緒に頑張りましょう！",
        "朝から元気にいきましょう！",
        "おはよう！今日も一日笑顔で過ごしましょう！",
        "朝の空気を感じて、一日をスタートしましょう！",
        "今日も頑張って！応援してるよ！",
        "一日がんばろう！私がついてるからね！",
        "自分にできることを一つずつこなしていきましょう！",
        "頑張れば必ず報われる！信じてるよ！",
        "失敗しても立ち上がって、次に繋げよう！",
        "一日中応援してるからね！がんばって！",
        "挑戦し続けることで成長できる！頑張って！",
        "今日も自分にチャレンジしよう！一緒に頑張ろう！",
        "前向きな気持ちで行こう！私が応援してるよ！",
        "自分の可能性を信じて、最後までやり抜こう！"
      ];
    }
    else if(state === "勤務中"){
      ans = ["すでに勤務中やで"]
    }
    else if(state === "休憩"){
      ans = [
        "ちょっとリフレッシュしてきてね。",
        "しっかり休憩して元気に戻ってきてね。",
        "休憩中にはゆっくり休んでね。",
        "また少しでもリラックスできるように頑張ってね。",
        "ゆっくりとした休憩を取って、リフレッシュしてね。",
        "ピッタリの休憩時間を過ごして、また頑張りましょう。",
        "しっかりと気分転換して、リフレッシュしてね。",
        "少しの時間でもリラックスできるといいね。",
        "休憩中はゆっくりと休んで、ストレスを解消しよう。",
        "しっかりと身体を休めて、また元気に仕事に戻ってきてね。",
        "休憩時間には、少しでもリフレッシュできるように頑張ってね。",
        "少しの休憩であっても、リラックスして充電しよう。",
        "しっかりと休憩を取って、また頑張りましょう。",
        "少しの時間でも、気分転換をしてリフレッシュしよう。",
        "休憩時間には、しっかりとリラックスして疲れをとってね。",
        "休憩中には、ゆっくりとした時間を過ごしてね。",
        "少しでもストレスを解消するために、休憩を利用してね。",
        "休憩中には、気分転換をしてリフレッシュするといいね。",
        "しっかりと身体を休めて、また頑張りましょう。",
        "少しの休憩でも、リラックスして心身ともにリフレッシュしよう。"
      ];
    }
    else if (state === "休憩中"){
      ans = ["すでに休憩中やで"]
    }
    else if(state === "再開"){
      ans = [
        "お帰りなさい！リフレッシュできましたか？",
        "休憩から戻ってきてくれてうれしいです。元気に戻れましたか？",
        "お疲れ様でした。休憩は十分に取れましたか？",
        "休憩から戻ってきたんですね。気分転換できたかな？",
        "休憩から帰ってきたようですね。リフレッシュしたかな？",
        "休憩から戻ってきたら、また頑張りましょう。",
        "休憩は十分に取れたみたいで、良かったですね。",
        "休憩から戻ってきてくれて、ありがとうございます。元気になれましたか？",
        "休憩から帰ってきてくれて、うれしいです。気分転換できましたか？",
        "休憩から戻ってきたようですね。身体は十分に休めましたか？",
        "お帰りなさい！休憩は効果的でしたか？",
        "休憩から戻ってきたんですね。また頑張りましょう。",
        "休憩から帰ってきてくれて、ありがとうございます。リフレッシュできましたか？",
        "休憩から戻ってきたようですね。疲れはとれましたか？",
        "休憩から帰ってきたんですね。気分転換できましたか？",
        "休憩から戻ってきてくれて、うれしいです。また頑張りましょう。",
        "休憩から帰ってきたようですね。また元気に仕事に取り組んでください。",
        "お帰りなさい！休憩は効果的でしたか？",
        "休憩から戻ってきたんですね。疲れはとれましたか？",
        "休憩から帰ってきてくれて、ありがとうございます。また一緒に頑張りましょう。"
      ]
    }
    else if (state === "勤務外"){
      ans = ["すでに退勤してるで"]
    }
    else if (state === "退勤"){
      ans = [
        "お疲れ様でした！お仕事お疲れ様！",
        "今日も一日、お疲れ様でした！",
        "ねえねえ、お仕事お疲れ様だよ！",
        "よく頑張りましたね！お疲れ様でした！",
        "お疲れ様！帰り道、気をつけてね！",
        "ほんとうに、お疲れ様でした！",
        "お疲れ様！明日もまた頑張ろうね！",
        "今日も一日お疲れ様でした！",
        "あらあら、もう終わっちゃったんだね！お疲れ様！",
        "うんうん、お仕事お疲れ様だよ！",
        "ねえねえ、お疲れ様って言ってあげようか？",
        "おつかれさまでした！帰り道、お気をつけて！",
        "あのね、お仕事お疲れ様！おうちに帰って、ゆっくり休んでね！",
        "ほんとにほんとに、お疲れ様でした！",
        "お疲れ様！明日もまたがんばろうね！",
        "今日も一日、よく頑張ったね！お疲れ様でした！",
        "あらあら、もうお終いなんだね！お疲れ様！",
        "うんうん、お疲れ様！今日も一日お疲れ様でした！",
        "ねえねえ、お仕事お疲れ様だよ！帰り道、気をつけてね！",
        "ほんとうに、よくがんばりましたね！お疲れ様でした！"
      ]
    }else{
      ans = [
        "えぇっと、すみません、もう一回言ってもらえますか？",
        "ごめんなさい、聞き逃しちゃったみたいで、もう一度お願いします！",
        "あら、もう一回言ってもらえると嬉しいわ！",
        "あのね、ちょっと言葉が難しいんだけど、もう一回言ってくれる？",
        "すみません、言い方がよくわからなかったので、もう一度説明していただけませんか？",
        "ごめんなさい、ちょっと混乱しちゃって、もう一度教えてもらっていい？",
        "あのね、もう一度聞かせてくれたら嬉しいんだけど、いい？",
        "えっと、もう一回言ってもらってもいいですか？すみません、ちょっとついていけてなくて…",
        "あら、もう一度言ってもらえると助かるわ。私、ちょっと遅れてしまったみたいで…",
        "すみません、言葉の意味がよくわからなかったんですけど、もう一回説明していただけませんか？",
        "ごめんなさい、もう一回言ってくれると嬉しいんです。ちょっと聞き逃してしまったみたいで…",
        "あのね、ちょっと言葉の意味がわからなくって…もう一度教えてくれる？",
        "えっと、もう一回言ってもらっていいですか？すみません、聞き取れなかったみたいで…",
        "ごめんなさい、もう一回言ってくれると助かるんです。ちょっと耳が遠いみたいで…",
        "あのね、もう一回言ってもらえると嬉しいんだけど、いいですか？ちょっと聞き逃しちゃったみたいで…",
        "すみません、もう一回言ってくれるとありがたいんです。ちょっと意味がわからなくって…",
        "あら、もう一回言ってもらえると嬉しいわ。ちょっと言葉のニュアンスがわからなくて…",
        "えっと、すみません、もう一回言ってくれると助かります。"
      ]
    }
    return { text: randomGet(ans) }
  }

  switch (command) {
    case '/kinchan':
    let response = { text: "空だとエラーになるからとりあえず入れてる文章やで" }
      if(text.includes("出勤")){
        if(userState === "勤務中"){
          response = setResponse("勤務中")
        }else if(userState === "休憩中"){
          setAwayTime("勤務中")
          postMessage("再開")
          response = setResponse("再開")
        }else if(userState === "勤務外"){
          setShukkin()
          postMessage("出勤")
          response = setResponse("出勤")
        }
      }else if(text.includes("休憩")){
        if(userState === "勤務中"){
          setWorkTime("休憩中")
          postMessage("休憩")
          response = setResponse("休憩")
        }else if(userState === "休憩中"){
          response = setResponse("休憩中")
        }else if(userState === "勤務外"){
          response = setResponse("勤務外")
        }
      }else if(text.includes("再開") || text.includes("戻った")){
        if(userState === "勤務中"){
          response = setResponse("勤務中")
        }else if(userState === "休憩中"){
          setAwayTime("勤務中")
          postMessage("再開")
          response = setResponse("再開")
        }else if(userState === "勤務外"){
          response = setResponse("勤務外")
        }
      }else if(text.includes("退勤")){
        if(userState === "勤務中"){
          setWorkTime("勤務外")
          setTaikin()
          postMessage("退勤")
          response = setResponse("退勤")
        }else if(userState === "休憩中"){
          setAwayTime("勤務外")
          setTaikin()
          postMessage("退勤")
          response = setResponse("休憩中")
        }else if(userState === "勤務外"){
          response = setResponse("勤務外")
        }
      }else{
        response = setResponse("例外")
      }
      return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
    default:
      return ContentService.createTextOutput(JSON.stringify(e));
  }
}

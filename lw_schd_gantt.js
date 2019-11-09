// エントリーポイント
function myFunction() {
  // 作業フォルダからSCHDが付いたファイルを取得
  var workfolder = DriveApp.getFoldersByName('SCHDGANTT_スケジュール').next();
  var files = workfolder.searchFiles('title contains "SCHD:"');
  // todayは今月の1日とする
  var today = new Date();
  today.setDate(1);
  Logger.log(today);

  // データ抽出
  var records = [];  
  while (files.hasNext()) {
    var file = files.next();
    // ファイルがスプレッドシートか確認
    if(file.getMimeType() === 'application/vnd.google-apps.spreadsheet'){
//      Logger.log(file.getName());
      var fname = file.getName();
      
      var spreadsheet = SpreadsheetApp.open(file);
      var sheet = spreadsheet.getSheetByName('スケジュール');
      if(sheet != null){
//        Logger.log(sheet.getName());
        var values = sheet.getDataRange().getValues();
        // 1行ずつ取り出す
        for(var r=1; r<values.length; r++){//GASでfor ofは動かない
          var task = values[r][0];
          var start = values[r][1];
          var end = values[r][2];
          var member = values[r][4];
//          Logger.log(task + start + end + member);
          // memberが＊のときはスキップ
          if(member == '＊' || member == '*') continue;
          // 今日の日付より終了日が前の場合もスキップ
          if(end < today) continue;
          // 複数メンバー対応　分割して別レコードにする
          var members = member.split(',');
          for(var m = 0; m < members.length; m++){
            records.push({
              'project': fname.slice(5),
              'task': task,
              'start': start, 
              'end': end,
              'member': members[m],
              'mcount': m
            });
          }
        }
      }
    }
  }
  // ガントチャート生成
  var svgtext = createGANTT(records, today);
  overwriteFile(workfolder, 'GANTT_' + today.getFullYear() +'_'+ ('0' + (today.getMonth()+1)).slice(-2) + '.svg', svgtext, 'image/svg+xml');
  var m2 = new Date(today.getFullYear(), today.getMonth()+1, 1);
  svgtext = createGANTT(records, m2);
  overwriteFile(workfolder, 'GANTT_' + m2.getFullYear() +'_'+ ('0' + (m2.getMonth()+1)).slice(-2) + '.svg', svgtext, 'image/svg+xml');
  var m3 = new Date(today.getFullYear(), today.getMonth()+2, 1);
  svgtext = createGANTT(records, m3);
  overwriteFile(workfolder, 'GANTT_' + m3.getFullYear() +'_'+ ('0' + (m3.getMonth()+1)).slice(-2) + '.svg', svgtext, 'image/svg+xml');
  
  //  overwriteFile(workfolder, 'test.txt', JSON.stringify(records), 'text/plain');
  // overwriteFile(workfolder, 'test.svg', '<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="100" height="60" fill="#ddd" /></svg>', 'image/svg+xml');
}

// ガントチャート生成
function createGANTT(records, startdate){
  // 月の末日を求める
  var enddate = new Date(startdate.getFullYear(), startdate.getMonth() + 1, 0);
  // 各種データの用意
  var svgtext = '<?xml version="1.0" encoding="utf-8"?><svg xmlns="http://www.w3.org/2000/svg" width="800" height="2000">';
  var ONEDAY = 20;  //1日を20pxとする
  svgtext += '<text x="0" y="16" font-size="16" fill="black">' 
  + startdate.getFullYear()  + '/' + parseInt(startdate.getMonth()+1) + '</text>';
  // 現在のレコード
  var curproject = '';
  //現在の基準位置
  var by = 0;
  var bx = 160;
  
  // 日付ラベル
  for(var x=0; x<=enddate.getDate(); x++){
      svgtext += '<rect x="' + (bx+parseInt((x+1)*ONEDAY)) + '" y="' + (by+20) 
      + '" width="' + ONEDAY + '" height="1000" stroke="cyan" fill="none"/>';    
      svgtext += '<text x="' + (bx+parseInt(x*ONEDAY)) + '" y="' + (by+34) + '" font-size="10" fill="black">' + x + '</text>';
  }
  
  for(var r=0; r<records.length; r++){
    // 終了時間がない場合はスキップ    
    if(records[r].end == '') continue;
    var ed = records[r].end.getDate();
    // 終了時間が期間の開始よりも前ならスキップ
    if(records[r].end < startdate) continue;
    // プロジェクト名が変わったときにbyをずらす
    if(records[r].project != curproject){
      curproject = records[r].project;
      by += 50;
      svgtext += '<text x="0" y="' + by + '" font-size="10" fill="black">' + curproject + '</text>';
      svgtext += '<line x1="0" y1="' + (by-12) + '" x2="1200" y2="' + (by-12) + '" stroke="blue"/>' ; 
    }
    // 開始時間がない
    if(records[r].start == ''){
      // 終了時間のみのときは線を引く
      // 終了時間がendateよりあとならスキップ
      if(records[r].end > enddate) continue;
      svgtext += '<line x1="' + (bx+parseInt(ed*ONEDAY)) + '" y1="' + (by-12) 
      + '" x2="' + (bx+parseInt(ed*ONEDAY)) + '" y2="' + (by+28) + '" stroke="red"/>';    
      svgtext += '<line x1="' + (bx+parseInt((ed-1)*ONEDAY)) + '" y1="' + (by+8) 
      + '" x2="' + (bx+parseInt(ed*ONEDAY)) + '" y2="' + (by+8) + '" stroke="red"/>';    
      svgtext += '<text x="' + (bx+parseInt(ed*ONEDAY)) + '" y="' + (by+records[r].mcount*10) + '" font-size="10" fill="black">' 
      + records[r].task + records[r].member + '</text>';
    } else {
      // 開始と終了がある
      var sd = records[r].start.getDate();
      // 期間の終了より先の予定ならスキップ
      if(records[r].start > enddate) continue;
      // 期間の範囲内に収める
      if(records[r].start < startdate) sd = startdate.getDate();
      if(records[r].end > enddate) ed = enddate.getDate();
      if(records[r].mcount == 0){
        svgtext += '<rect x="' + (bx+parseInt((sd-1)*ONEDAY)) + '" y="' + (by-12) 
        + '" width="' + (parseInt((ed-sd+1)*ONEDAY)) + '" height="40" stroke="red" fill="pink" fill-opacity="0.5"/>';    
      }
      svgtext += '<text x="' + (bx+parseInt((sd-1)*ONEDAY)) + '" y="' + (by+records[r].mcount*10) + '" font-size="10" fill="black">' 
      + records[r].task + records[r].member + '</text>';
    }
  }
  svgtext += '</svg>';
  return svgtext;
}


// ファイルの上書き（同名ファイルを消去してから保存）
function overwriteFile(workfolder, name, content, mimetype){
  var files = workfolder.getFilesByName(name);
  if(files.hasNext()){
    workfolder.removeFile(files.next());
  }
  workfolder.createFile(name, content, mimetype);  
}
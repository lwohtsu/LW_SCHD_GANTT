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
              'project': escapeHtml(fname.slice(5)),
              'task': escapeHtml(task),
              'start': start, 
              'end': end,
              'member': escapeHtml(members[m])
            });
          }
        }
      }
    }
  }
  // ガントチャート生成
  var ganttfolder = workfolder.getFoldersByName('GANTT').next();
  for(var i=0; i<6; i++){
    var m = new Date(today.getFullYear(), today.getMonth()+i, 1);
    svgtext = createGANTT(records, m);
    overwriteFile(ganttfolder, 'GANTT_' + m.getFullYear() +'_'+ ('0' + (m.getMonth()+1)).slice(-2) + '.svg', svgtext, 'image/svg+xml');
  }
  // 担当者別ガントチャート生成
  records.sort(function(a,b){
    if (a.member < b.member) return -1;
    if (a.member > b.member) return 1;
    if (a.project < b.project) return -1;
    if (a.project > b.project) return 1;
    return 0;
  });
//  Logger.log(records);
  for(var i=0; i<6; i++){
    var m = new Date(today.getFullYear(), today.getMonth()+i, 1);
    svgtext = createGANTTbyAssign(records, m);
    overwriteFile(ganttfolder, 'ASSIGN_' + m.getFullYear() +'_'+ ('0' + (m.getMonth()+1)).slice(-2) + '.svg', svgtext, 'image/svg+xml');
  }
  //  overwriteFile(workfolder, 'test.txt', JSON.stringify(records), 'text/plain');
  // overwriteFile(workfolder, 'test.svg', '<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="100" height="60" fill="#ddd" /></svg>', 'image/svg+xml');
}

// ガントチャート生成
function createGANTT(records, startdate){
  // 月の末日を求める
  var enddate = new Date(startdate.getFullYear(), startdate.getMonth() + 1, 0);
  // 各種データの用意
  var svgtext = '';
  var ONEDAY = 20;  //1日を20pxとする
  svgtext += '<text x="0" y="16" font-size="16" fill="black">' 
  + startdate.getFullYear()  + '/' + parseInt(startdate.getMonth()+1) + '</text>';
  // 現在のプロジェクト
  var curproject = '';
  //現在の基準位置
  var by = 0;
  var bx = 160;
  var bheight = 50;
  // タスクの重なりを防止する
  var overmap = [];  
  
  // 日付ラベル
  for(var x=1; x<=enddate.getDate(); x++){
      svgtext += '<rect x="' + (bx+parseInt(x*ONEDAY)) + '" y="' + (by+20) 
      + '" width="' + ONEDAY + '" height="1000" stroke="cyan" fill="none"/>';    
      svgtext += '<text x="' + (bx+parseInt(x*ONEDAY)) + '" y="' + (by+34) + '" font-size="10" fill="black">' + x + '</text>';
  }

  // ループで全レコードを処理
  for(var r=0; r<records.length; r++){
    // 終了時間がない場合はスキップ    
    if(records[r].end == '') continue;
    var ed = records[r].end.getDate();
    // 終了時間が期間の開始よりも前ならスキップ
    if(records[r].end < startdate) continue;
    // 開始時間が期間の終了よりも後ならスキップ
    if(records[r].start > enddate) continue;
    // プロジェクト名が変わったときにbyをずらす
    if(records[r].project != curproject){
      curproject = records[r].project;
      by += bheight;
      bheight = 50;
      svgtext += '<text x="0" y="' + by + '" font-size="10" fill="black">' + curproject + '</text>';
      svgtext += '<line x1="0" y1="' + (by-12) + '" x2="1200" y2="' + (by-12) + '" stroke="blue"/>' ; 
      overmap = [];
    }
    // 開始時間がない
    if(records[r].start == ''){
      // 終了時間のみのときは線を引く
      // 終了時間がendateよりあとならスキップ
      if(records[r].end > enddate) continue;
      // 重なりチェック
      var yshift = 0;
      for (var i=0; i < overmap.length; i++){
        if(overmap[i].sd <= ed && overmap[i].ed >= ed){
          if(overmap[i].yshift == yshift){
            yshift++;
          }
        }
      }
      if((yshift+1)*11 > bheight){
        bheight += 11;
      }
      // 描画
      svgtext += '<line x1="' + (bx+parseInt((ed+1)*ONEDAY)) + '" y1="' + (by-12) 
      + '" x2="' + (bx+parseInt((ed+1)*ONEDAY)) + '" y2="' + (by+28) + '" stroke="red"/>';    
      svgtext += '<line x1="' + (bx+parseInt(ed*ONEDAY)) + '" y1="' + (by+8) 
      + '" x2="' + (bx+parseInt((ed+1)*ONEDAY)) + '" y2="' + (by+8) + '" stroke="red"/>';    
      svgtext += '<text x="' + (bx+parseInt((ed+1)*ONEDAY)) + '" y="' + (by+yshift*11-1) + '" font-size="10" fill="black">' 
      + records[r].task + records[r].member + '</text>';
      overmap.push({"sd": ed, "ed": ed, "yshift": yshift});
    } else {
      // 開始と終了がある
      var sd = records[r].start.getDate();
      // 期間の終了より先の予定ならスキップ
      if(records[r].start > enddate) continue;
      // 期間の範囲内に収める
      if(records[r].start < startdate) sd = startdate.getDate();
      if(records[r].end > enddate) ed = enddate.getDate();
      // 重なりチェック
      var yshift = 0;
      for (var i=0; i < overmap.length; i++){
        if(overmap[i].sd <= ed && overmap[i].ed >= sd){
          if(overmap[i].yshift == yshift){
            yshift++;
          }
        }
      }
      if((yshift+1)*11 > bheight){
        bheight += 11;
      }
      // 描画
      if(records[r].task.indexOf('_')==0){
        svgtext += '<rect x="' + (bx+parseInt(sd*ONEDAY)) + '" y="' + (by+yshift*11-11) 
        + '" width="' + (parseInt((ed-sd+1)*ONEDAY)) + '" height="11" stroke="red" fill="gray" fill-opacity="0.5" stroke-dasharray="2,5"/>';    
      } else {
        svgtext += '<rect x="' + (bx+parseInt(sd*ONEDAY)) + '" y="' + (by+yshift*11-11) 
        + '" width="' + (parseInt((ed-sd+1)*ONEDAY)) + '" height="11" stroke="red" fill="pink" fill-opacity="0.5"/>';    
      }
      svgtext += '<text x="' + (bx+parseInt(sd*ONEDAY)) + '" y="' + (by+yshift*11-1) + '" font-size="10" fill="black">' 
      + records[r].task + records[r].member + '</text>';
      overmap.push({"sd": sd, "ed": ed, "yshift": yshift});
    }
  }
  svgtext = '<?xml version="1.0" encoding="utf-8"?><svg xmlns="http://www.w3.org/2000/svg" width="900" height="' + (by+60) + '">' + svgtext + '</svg>';
  return svgtext;
}

// 人別のガントチャート生成
function createGANTTbyAssign(records, startdate){
  // 月の末日を求める
  var enddate = new Date(startdate.getFullYear(), startdate.getMonth() + 1, 0);
  // 各種データの用意
  var svgtext = '';
  var ONEDAY = 20;  //1日を20pxとする
  svgtext += '<text x="0" y="16" font-size="16" fill="black">' 
  + startdate.getFullYear()  + '/' + parseInt(startdate.getMonth()+1) + '</text>';
  // 現在の担当者
  var curproject = '';
  var curmember = '_____';
  //現在の基準位置
  var by = 0;
  var bx = 240;
  var bheight = 20;
  // タスクの重なりを防止する
  var overmap = [];  
  
  // 日付ラベル
  for(var x=1; x<=enddate.getDate(); x++){
      svgtext += '<rect x="' + (bx+parseInt(x*ONEDAY)) + '" y="' + (by+20) 
      + '" width="' + ONEDAY + '" height="1000" stroke="cyan" fill="none"/>';    
      svgtext += '<text x="' + (bx+parseInt(x*ONEDAY)) + '" y="' + (by+34) + '" font-size="10" fill="black">' + x + '</text>';
  }
  by = 30;

  // ループで全レコードを処理
  for(var r=0; r<records.length; r++){
    // 終了時間がない場合はスキップ    
    if(records[r].end == '') continue;
    // 担当者未記入もスキップ
    if(records[r].member == '') continue;
    var ed = records[r].end.getDate();
    // 終了時間が期間の開始よりも前ならスキップ
    if(records[r].end < startdate) continue;
    // 開始時間が期間の終了よりも後ならスキップ
    if(records[r].start > enddate) continue;
    // 担当者名かプロジェクト名が変わったときにbyをずらす
    if(records[r].member != curmember){
      curmember = records[r].member;
      curproject = records[r].project;
      by += bheight;
      bheight = 20;
      svgtext += '<text x="0" y="' + by + '" font-size="10" fill="black">' + curmember +':' + curproject + '</text>';
      svgtext += '<line x1="0" y1="' + (by-12) + '" x2="1200" y2="' + (by-12) + '" stroke="blue"/>' ; 
      overmap = [];
    } else if(records[r].project != curproject) {
      curproject = records[r].project;
      by += bheight;
      bheight = 20;
      svgtext += '<text x="0" y="' + by + '" font-size="10" fill="black">' + curmember +':' + curproject + '</text>';
      svgtext += '<line x1="20" y1="' + (by-12) + '" x2="1200" y2="' + (by-12) + '" stroke="blue" stroke-dasharray="2,5"/>' ; 
      overmap = [];      
    }
    // 開始時間がない
    if(records[r].start == ''){
      // 終了時間のみのときは線を引く
      // 終了時間がendateよりあとならスキップ
      if(records[r].end > enddate) continue;
      // 重なりチェック
      var yshift = 0;
      for (var i=0; i < overmap.length; i++){
        if(overmap[i].sd <= ed && overmap[i].ed >= ed){
          if(overmap[i].yshift == yshift){
            yshift++;
          }
        }
      }
      if((yshift+1)*11 > bheight){
        bheight += 11;
      }
      // 描画
      svgtext += '<line x1="' + (bx+parseInt((ed+1)*ONEDAY)) + '" y1="' + (by-12) 
      + '" x2="' + (bx+parseInt((ed+1)*ONEDAY)) + '" y2="' + (by) + '" stroke="red"/>';    
      svgtext += '<line x1="' + (bx+parseInt(ed*ONEDAY)) + '" y1="' + (by-6) 
      + '" x2="' + (bx+parseInt((ed+1)*ONEDAY)) + '" y2="' + (by-6) + '" stroke="red"/>';    
      svgtext += '<text x="' + (bx+parseInt((ed+1)*ONEDAY)) + '" y="' + (by+yshift*11-1) + '" font-size="10" fill="black">' 
       + records[r].task + '</text>';
      overmap.push({"sd": ed, "ed": ed, "yshift": yshift});
    } else {
      // 開始と終了がある
      var sd = records[r].start.getDate();
      // 期間の終了より先の予定ならスキップ
      if(records[r].start > enddate) continue;
      // 期間の範囲内に収める
      if(records[r].start < startdate) sd = startdate.getDate();
      if(records[r].end > enddate) ed = enddate.getDate();
      // 重なりチェック
      var yshift = 0;
      for (var i=0; i < overmap.length; i++){
        if(overmap[i].sd <= ed && overmap[i].ed >= sd){
          if(overmap[i].yshift == yshift){
            yshift++;
          }
        }
      }
      if((yshift+1)*11 > bheight){
        bheight += 11;
      }
      // 描画
      if(records[r].task.indexOf('_')==0){
        svgtext += '<rect x="' + (bx+parseInt(sd*ONEDAY)) + '" y="' + (by+yshift*11-11) 
        + '" width="' + (parseInt((ed-sd+1)*ONEDAY)) + '" height="11" stroke="red" fill="gray" fill-opacity="0.5" stroke-dasharray="2,5"/>';    
      } else {
        svgtext += '<rect x="' + (bx+parseInt(sd*ONEDAY)) + '" y="' + (by+yshift*11-11) 
        + '" width="' + (parseInt((ed-sd+1)*ONEDAY)) + '" height="11" stroke="red" fill="pink" fill-opacity="0.5"/>';    
      }
      svgtext += '<text x="' + (bx+parseInt(sd*ONEDAY)) + '" y="' + (by+yshift*11-1) + '" font-size="10" fill="black">' 
      + records[r].task + '</text>';
      overmap.push({"sd": sd, "ed": ed, "yshift": yshift});
    }
  }
  svgtext = '<?xml version="1.0" encoding="utf-8"?><svg xmlns="http://www.w3.org/2000/svg" width="900" height="' + (by+60) + '">' + svgtext + '</svg>';
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

// エスケープ処理
function escapeHtml(unsafe) {
  return unsafe
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;")
  .replace(/'/g, "&#039;");
}
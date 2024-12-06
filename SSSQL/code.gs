const j01="SQLクエリでスプレッドシートのデータ抽出ができます";
const j02="<font color=\"#043c78\">シート名がテーブル名になります<br>シート内の1行目は列名（見出し）で2行目以降からがデータです</font>";
const j03="<font color=\"#0000ff\">[</font> <b>SELECT構文</b> <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>SELECT</b></font> ( <font color=\"#191970\">DISTINCT</font> )<br>&nbsp;&nbsp; ( 集計関数 ) * | ( シート名. ) 列名 , ... <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>FROM</b></font> シート名<br>&nbsp;&nbsp; <font color=\"#c71585\">[</font> <font color=\"#191970\"><b>INNER</b></font><br>&nbsp;&nbsp; | ( <font color=\"#191970\">LEFT</font> | <font color=\"#191970\">RIGHT</font> | <font color=\"#191970\">FULL</font> ) <font color=\"#191970\"><b>OUTER</b></font><br>&nbsp;&nbsp; | <font color=\"#191970\"><b>CROSS</b></font><br>&nbsp;&nbsp;&nbsp;&nbsp; <font color=\"#191970\"><b>JOIN</b></font> シート名 <font color=\"#191970\"><b>ON</b></font> 結合条件 <font color=\"#c71585\">]</font> <font color=\"#c71585\">...</font> <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>WHERE</b></font> 検索条件 <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>GROUP BY</b></font> ( シート名. ) 列名 , ...<br>&nbsp;&nbsp; <font color=\"#c71585\">[</font> <font color=\"#191970\"><b>HAVING</b></font> 検索条件 <font color=\"#c71585\">]</font> <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>ORDER BY</b></font> ( シート名. ) 列名<br>&nbsp;&nbsp; ( <font color=\"#191970\">ASC</font> | <font color=\"#191970\">DESC</font> ) , ... <font color=\"#0000ff\">]</font> <br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>LIMIT</b></font> 上限レコード数 <font color=\"#0000ff\">]</font><br><font color=\"#f09199\">サブクエリ（副問合せ）は非対応です</font>";
const j04="SQLクエリ";
const j05="SQLクエリ実行結果出力先";
const j06="ノーティフィケーション";
const j07="アクティブシート";
const j08="CSVファイル";
const j09="クリップボード";
const j10="SQLクエリ実行";
const j11="ノーティフィケーション ▼";
const j12="クリア";
const j13="※ エラーが発生しました ※";
const j14="SQLクエリをご確認ください";
const j15="シートの書込み権限をご確認ください";
const j16="この機能は編集権限があるシート上でないと利用できません";
const j17="データ件数：";
const j18="レコード";
const j19="アクティブシートに出力";
const j20="ダウンロードが完了しました";
const j21="CSVファイル";
const j22="CSVファイルに出力";
const j23="クリップボードにコピーしました";
const j24="クリップボード";
const j25="クリップボードに出力";
const j26="SELECT句がありません\n";
const j27="FROM句がありません\n";
const j28="のシートが存在しません\n";
const j29=" に ";
const j30=" が見つかりません";
const j31="SQL履歴";
const j32="SSSQL履歴";
const j33="SSSQL（SpreadSheetSQL）";
const j34="SELECT * FROM シート1\nWHERE 名前 LIKE '%佑藤%'\n\n\n\n";
const e01="Extract data from SpreadSheet using SQL Query";
const e02="<font color=\"#043c78\">Sheet name is table name.<br>The first row in the sheet is the column name (heading), and the second and subsequent rows are the data.</font>";
const e03="<font color=\"#0000ff\">[</font> <b>SELECT statement</b> <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>SELECT</b></font> ( <font color=\"#191970\">DISTINCT</font> )<br>&nbsp;&nbsp; ( Aggregate function ) * | ( Sheet name. ) column name , ... <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>FROM</b></font> Sheet name<br>&nbsp;&nbsp; <font color=\"#c71585\">[</font> <font color=\"#191970\"><b>INNER</b></font><br>&nbsp;&nbsp; | ( <font color=\"#191970\">LEFT</font> | <font color=\"#191970\">RIGHT</font> | <font color=\"#191970\">FULL</font> ) <font color=\"#191970\"><b>OUTER</b></font><br>&nbsp;&nbsp; | <font color=\"#191970\"><b>CROSS</b></font><br>&nbsp;&nbsp;&nbsp;&nbsp; <font color=\"#191970\"><b>JOIN</b></font> Sheet name <font color=\"#191970\"><b>ON</b></font> Join conditions <font color=\"#c71585\">]</font> <font color=\"#c71585\">...</font> <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>WHERE</b></font> Search conditions <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>GROUP BY</b></font> ( Sheet name. ) column name , ...<br>&nbsp;&nbsp; <font color=\"#c71585\">[</font> <font color=\"#191970\"><b>HAVING</b></font> Search conditions <font color=\"#c71585\">]</font> <font color=\"#0000ff\">]</font><br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>ORDER BY</b></font> ( Sheet name. ) column name<br>&nbsp;&nbsp; ( <font color=\"#191970\">ASC</font> | <font color=\"#191970\">DESC</font> ) , ... <font color=\"#0000ff\">]</font> <br><font color=\"#0000ff\">[</font> <font color=\"#191970\"><b>LIMIT</b></font> Limit records <font color=\"#0000ff\">]</font><br><font color=\"#f09199\">SubQuery is not supported.</font>";
const e04="SQL Query";
const e05="SQL Query result output";
const e06="Notification";
const e07="ActiveSheet";
const e08="CSV file";
const e09="Clipboard";
const e10="SQL Query execution";
const e11="Notification ▼";
const e12="Clear";
const e13="An error has occurred";
const e14="Please check your SQL query again.";
const e15="Please check the write permissions on the sheet.";
const e16="This feature is only available on sheets you have editing permissions to.";
const e17="Number:";
const e18="Record";
const e19="Output to ActiveSheet";
const e20="Download completed";
const e21="CSV file";
const e22="Output to CSV file";
const e23="Copied to Clipboard";
const e24="Clipboard";
const e25="Output to Clipboard";
const e26="SELECT statement is missing\n";
const e27="FROM statement is missing\n";
const e28="sheet does not exist\n";
const e29=" to ";
const e30=" not found";
const e31="History";
const e32="SSSQL History";
const e33="SSSQL(SpreadSheetSQL)";
const e34="SELECT * FROM SHEET1\nWHERE NAME IS NOT NULL\n\n\n\n";
var m01=j01;
var m02=j02;
var m03=j03;
var m04=j04;
var m05=j05;
var m06=j06;
var m07=j07;
var m08=j08;
var m09=j09;
var m10=j10;
var m11=j11;
var m12=j12;
var m13=j13;
var m14=j14;
var m15=j15;
var m16=j16;
var m17=j17;
var m18=j18;
var m19=j19;
var m20=j20;
var m21=j21;
var m22=j22;
var m23=j23;
var m24=j24;
var m25=j25;
var m26=j26;
var m27=j27;
var m28=j28;
var m29=j29;
var m30=j30;
var m31=j31;
var m32=j32;
var m33=j33;
var m34=j34;
//------------------------------------------------------------------------------------------------------------
//アドオン実行トリガー
function onDefOpen() { 
  if (Session.getActiveUserLocale() != 'ja') { 
    m01=e01;
    m02=e02;
    m03=e03;
    m04=e04;
    m05=e05;
    m06=e06;
    m07=e07;
    m08=e08;
    m09=e09;
    m10=e10;
    m11=e11;
    m12=e12;
    m13=e13;
    m14=e14;
    m15=e15;
    m16=e16;
    m17=e17;
    m18=e18;
    m19=e19;
    m20=e20;
    m21=e21;
    m22=e22;
    m23=e23;
    m24=e24;
    m25=e25;
    m26=e26;
    m27=e27;
    m28=e28;
    m29=e29;
    m30=e30;
    m31=e31;
    m32=e32;
    m33=e33;
    m34=e34;
  } 
  return buildCard(); 
}
//------------------------------------------------------------------------------------------------------------
//カードベースUI作成
function buildCard(response=m33, sqltext=m34, drcheck="c1", ntext=m03) {
  let bgtext = PropertiesService.getUserProperties().getProperty('SQL');
  if (bgtext && sqltext == m34) { sqltext=bgtext; }

  let cardHeader = CardService.newCardHeader().setSubtitle(m01).setImageUrl('https://raw.githubusercontent.com/google/material-design-icons/master/png/editor/query_stats/materialiconstwotone/48dp/2x/twotone_query_stats_black_48dp.png').setImageStyle(CardService.ImageStyle.CIRCLE);
  let section = CardService.newCardSection();
  section.addWidget(CardService.newDecoratedText().setBottomLabel(m02).setWrapText(true));
  section.addWidget(CardService.newTextInput().setFieldName('SSSQL').setTitle(m04).setValue(sqltext).setMultiline(true));
  section.addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.DROPDOWN).setTitle(m05).setFieldName("CHECKF").addItem(m06, "c1", (drcheck=="c1")).addItem(m07, "c2", (drcheck=="c2")).addItem(m08, "c3", (drcheck=="c3")).addItem(m09, "c4", (drcheck=="c4")));
  section.addWidget(CardService.newTextButton().setText(m10).setOnClickAction(CardService.newAction().setFunctionName('exeOnClick')).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  section.addWidget(CardService.newDecoratedText().setTopLabel(m11).setText(response).setBottomLabel(ntext).setWrapText(true));
  let fixedFooter = CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText(m12).setOnClickAction(CardService.newAction().setFunctionName("exeClear"))).setSecondaryButton(CardService.newTextButton().setText(m31).setOnClickAction(CardService.newAction().setFunctionName("exeHis")));
  
  return CardService.newCardBuilder().setHeader(cardHeader).setFixedFooter(fixedFooter).addSection(section).build();
}
//------------------------------------------------------------------------------------------------------------
//実行結果出力先をクリア
function exeClear(event) {
  let ss = event.formInput['SSSQL']; 
  let cf = event.formInput['CHECKF'];
  
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(buildCard("",ss,cf,""))).setStateChanged(true).build();
}
//------------------------------------------------------------------------------------------------------------
//SSSQLクエリ履歴
function exeHis(event) {  
  try {
    const setlog = new Set();
    let jsnlog = PropertiesService.getUserProperties().getProperty('HIS');
    if (jsnlog) {
      let hoge = JSON.parse(jsnlog);
      for (const geho of hoge) { if (!setlog.has(geho)) { setlog.add(geho); } }
    }
    let arr = Array.from(setlog);
    let hoge = '<hr>';
    for (let i = 0; i < setlog.size; i++) {
      let result = arr.pop();
      hoge = hoge + '<table><tbody><tr><td>'+result.replaceAll('\n','<br>')+'</td></tr></tbody></table><hr>';
    }
    let contents = '<html><head><base target="_top"></head><body>' + hoge + '</body></html>';
    const html = HtmlService.createHtmlOutput(contents);
    SpreadsheetApp.getUi().showModelessDialog(html.setWidth(450).setHeight(800), m32);
  } catch(e) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(buildCard("<font color=\"#ff0000\">"+m13+"</font>",event.formInput['SSSQL'],event.formInput['CHECKF'],m16+"<br>"+e.message))).setStateChanged(true).build();
  }
}
//------------------------------------------------------------------------------------------------------------
//SQLクエリ実行
function exeOnClick(event) {
  let ss = event.formInput['SSSQL']; 
  let cf = event.formInput['CHECKF'];
  let outarray = [];
  let ntext = "";
  let tes = "";
  const wiz = ss;
  const setlog = new Set();

  try {
    outarray = Query_res(ss);
    PropertiesService.getUserProperties().setProperty('SQL',wiz);
    let jsnlog = PropertiesService.getUserProperties().getProperty('HIS');
    if (jsnlog) {
      let hoge = JSON.parse(jsnlog);
      for (const geho of hoge) { if (!setlog.has(geho)) { setlog.add(geho); } }
    }
    if (!setlog.has(wiz)) { 
      setlog.add(wiz);
      let tarr = Array.from(setlog);
      if (setlog.size > 10) { tarr.shift(); }
      PropertiesService.getUserProperties().setProperty('HIS',JSON.stringify(tarr));
    }
  } catch(e) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(buildCard("<font color=\"#ff0000\">"+m13+"</font>",wiz,cf,m14+"<br>"+e.message))).setStateChanged(true).build();
  }
  ntext = "<font color=\"#043c78\">"+m17+(outarray.length-1)+m18+"</font>";
  if (cf == "c2") {
    try {
      const asheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      asheet.clear();
      asheet.setFrozenRows(1);  
      asheet.getRange(1,1,outarray.length,outarray[0].length).setNumberFormat('@');
      asheet.getRange(1,1,outarray.length,outarray[0].length).setValues(outarray);
      tes = m19;
    } catch(e) {
      return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(buildCard("<font color=\"#ff0000\">"+m13+"</font>",wiz,cf,m15+"<br>"+e.message))).setStateChanged(true).build();
    }
  } else if (cf == "c3" || cf == "c4") {
    for (const array of outarray) { tes = tes + array + "\n"; }
    const html = HtmlService.createTemplateFromFile('index');
    html.content = tes;
    html.checkf = cf;
    let fname = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() + Utilities.formatDate( new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '.csv'; 
    let ms1 = m20+"<br>"+fname;
    let ms2 = m21;
    tes = m22;
    if (cf == "c4") {
      ms1 = m23;
      ms2 = m24;
      tes = m25;
    }
    html.filename = fname;
    html.msg = ms1;
    try {
      SpreadsheetApp.getUi().showModelessDialog(html.evaluate().setWidth(320).setHeight(120), ms2);
    } catch(e) {
      return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(buildCard("<font color=\"#ff0000\">"+m13+"</font>",wiz,cf,m16+"<br>"+e.message))).setStateChanged(true).build();
    }
  } else if (cf == "c1"){
    for (const array of outarray) { tes = tes + array + "<br>"; }
    tes = "<font color=\"#191970\">" + tes + "</font>";
  }
  
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(buildCard(ntext,wiz,cf,tes))).setStateChanged(true).build();
}
//------------------------------------------------------------------------------------------------------------
//SQLクエリ解析
function Query_res(ss) {
  ss = ss.replaceAll('\n',' ').replaceAll(';','');
  if (ss.search(/SELECT/i) == -1) {
    throw new Error(m26);
  }
  if (ss.search(/FROM/i) == -1) {
    throw new Error(m27);
  }
  //SELECT句
  let select = ss.substring(ss.search(/SELECT/i)+6,ss.search(/FROM/i)).trim();
  let dis_fg = false; //distinct
  if (select.search(/DISTINCT/i) != -1) {
    select = select.substring(8).trim();
    dis_fg = true;
  }
  //FROM句
  let from = ss.substring(ss.search(/FROM/i)+4).trim();
  const tsql = ['WHERE', 'GROUP', 'ORDER', 'LIMIT'];
  for (const wsql in tsql) {
    let treg = new RegExp(tsql[wsql], "i");
    if (ss.search(treg) != -1) {
      from = ss.substring(ss.search(/FROM/i)+4, ss.search(treg)).trim();
      break;
    }
  }
  //WHERE句
  let where = "";
  if (ss.search(/WHERE/i) != -1) { 
    where = ss.substring(ss.search(/WHERE/i)+5).trim();
    const usql = ['GROUP', 'ORDER', 'LIMIT'];
    for (const vsql in usql) {
      let ureg = new RegExp(usql[vsql], "i");
      if (ss.search(ureg) != -1) {
        where = ss.substring(ss.search(/WHERE/i)+5, ss.search(ureg)).trim();
        break;
      }
    }
  }
  //WHERE句のBETWEENとINの対応
  where = sql_to_script(where);

  //GROUP BY句(HAVING句)
  let group = "";
  let having = "";
  if (ss.search(/GROUP/i) != -1) {
    group = ss.substring(ss.search(/GROUP/i)).trim();
    if (ss.search(/HAVING/i) != -1) {
      group = ss.substring(ss.search(/GROUP/i),ss.search(/HAVING/i)).trim();
      having = ss.substring(ss.search(/HAVING/i)+6).trim();
      if (ss.search(/ORDER/i) != -1) {
        having = ss.substring(ss.search(/HAVING/i)+6,ss.search(/ORDER/i)).trim();
      } else if (ss.search(/LIMIT/i) != -1) {
        having = ss.substring(ss.search(/HAVING/i)+6,ss.search(/LIMIT/i)).trim();
      }
    } else if (ss.search(/ORDER/i) != -1) { 
      group = ss.substring(ss.search(/GROUP/i),ss.search(/ORDER/i)).trim();
    } else if (ss.search(/LIMIT/i) != -1) {
      group = ss.substring(ss.search(/GROUP/i),ss.search(/LIMIT/i)).trim();
    }
  }
  //ORDER BY句
  let order = "";
  if (ss.search(/ORDER/i) != -1) { 
    order = ss.substring(ss.search(/ORDER/i)).trim();
    if (ss.search(/LIMIT/i) != -1) {
      order = ss.substring(ss.search(/ORDER/i), ss.search(/LIMIT/i)).trim();
    }
  }

  let tables = []; //0:テーブル名, 1:全列名, 2:JOIN種類, 3:JOIN条件, 4:シートのデータ範囲の文字列, 5:見出し, 6:Query関数, 7:SH_ID
  let from1 = from;
  let ct=1; 
  while(ct) {
    ct=0;
    let jos = "";
    let jon = "";
    let from0=from1;
    const joins = ['LEFT' , 'RIGHT' , 'INNER' , 'FULL' , 'CROSS', 'JOIN'];
    for (const join of joins) {
      let reg = new RegExp(join, "i");
      if (from1.search(reg) != -1) {        
        from0 = from1.substring(0, from1.search(reg)).trim();
        from1 = from1.substring(from1.search(/JOIN/i)+4).trim();
        jos = join;
        ct=1;
        break;
      }
    }
    if (from0.search(/ON/i) != -1) {
      jon = from0.substring(from0.search(/ON/i)+2).trim();
      from0 = from0.substring(0, from0.search(/ON/i)).trim();
    }
    tables.push([from0,'',jos,jon,'','','SELECT ',0]);
  }

  //tables[0：テーブル名][1：全列名|4：シートのデータ範囲の文字列|7：SH_ID]
  const hakua = /, {1,}/ig;
  const hakub = / {1,},/ig;
  for (let ct = 0; ct < tables.length; ct++) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tables[ct][0]);
    if (!sheet) { throw new Error(tables[ct][0]+m28); }
    let lasCol = sheet.getRange(1, sheet.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();
    if (lasCol == 1) { lasCol = sheet.getLastColumn(); }
    let t_table  = sheet.getRange(1, 1, 1, lasCol).getValues();
    tables[ct][1] = t_table[0].join().replaceAll('\n','').replaceAll(hakua,',').replaceAll(hakub,',').trim();
    tables[ct][4] = sheet.getRange(2, 1, sheet.getLastRow()-1, lasCol).getA1Notation();
    tables[ct][7] = sheet.getSheetId();
  }

  let outarray = []; //最終出力用の配列
  select = select.replaceAll(hakua,',').replaceAll(hakub,',');
  if (tables.length == 1) { //テーブル結合なし
    if (select == "*") {
      tables[0][5] = tables[0][1];
    } else {
      tables[0][5] = select;
    }
    select = tables[0][5];
    if (ss.search(/LIMIT/i) != -1) {
      order = order + " " + ss.substring(ss.search(/LIMIT/i));
    }
    outarray = outarray.concat(getd(tables,select,where,group,having,order));
    outarray.unshift(select.split(","));
  } else { //テーブル結合あり
    outarray = outarray.concat(getds(tables,select,where,group,having,order));
  }

  //SELECT(DISTINCT処理)
  if (dis_fg) { 
    outarray = [...new Set(outarray.map(JSON.stringify))].map(JSON.parse); 
  }
  //LIMIT処理
  if (ss.search(/LIMIT/i) != -1 && tables.length > 1) {
    let tlimit = parseInt(ss.substring(ss.search(/LIMIT/i)+5).trim(),10);
    if (tlimit != "NaN") { outarray.splice((tlimit+1), (outarray.length-tlimit-1)); }
  }

  return outarray;
}
//------------------------------------------------------------------------------------------------------------
//テーブル結合がないクエリ用
function getd(tables,select,where,group,having,order) {
  const BASE_URL = "https://docs.google.com/spreadsheets/d/";
  const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const token = ScriptApp.getOAuthToken();
  const headers = {
    "Authorization": "Bearer " + token
  };
  const options = {
    "contentType": "application/json",
    "headers": headers,
    "muteHttpExceptions": true
  };
  if (where != "") { where = " WHERE "+where; }
  let query = "SELECT " + select + where + " " + group + " " + order;
  let tmp_cols = tables[0][1].split(",");
  let tmp_h = [];
  for (let key in tmp_cols){
    tmp_h.push({ no: ConvertToLetter(key), name: tmp_cols[key]});
  }
  tmp_h.sort((a, b) => a.name.length > b.name.length ? -1 : 1);
  for (let key in tmp_h){
    query = query.replaceAll(tmp_h[key].name,tmp_h[key].no); //クエリの列名を英数字に変換
  }
  //Google Visualization API
  let URL = BASE_URL + SS_ID + "/gviz/tq?gid=" + tables[0][7] + "&range=" + tables[0][4] + "&tqx=out:json&tq=";
  let response = UrlFetchApp.fetch(URL + encodeURIComponent(query.trim()), options);
  let jobj = get_jobj(response);
  
  //HAVING句の独自処理
  if (having) {
    having = sql_to_script(having);
    having = having.replaceAll('=','==').replaceAll('!==','!=').replaceAll('>==','>=').replaceAll('<==','<=').replaceAll(/ AND /ig,' && ').replaceAll(/ OR /ig,' || ');
    having = having.replaceAll(/IS {1,}NULL/ig,"== ''").replaceAll(/IS {1,}NOT {1,}NULL/ig,"!= ''");
    having = like_to_script(having);
    let select_a = select.split(",");
    select_a.sort(function(a, b) {return b.length - a.length;});
    const select_b = select.split(",");
    select_a.forEach((value) => {
      having = having.replaceAll(value,'wiz['+select_b.findIndex(select_v => select_v == value)+']');
    });
    return get_array(jobj).filter(function(wiz){ return eval(having); });
  }
  
  return get_array(jobj);
}
//------------------------------------------------------------------------------------------------------------
//テーブル結合があるクエリ用
function getds(tables,select,where,group,having,order) {
  const BASE_URL = "https://docs.google.com/spreadsheets/d/";
  const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const token = ScriptApp.getOAuthToken();
  const headers = {
    "Authorization": "Bearer " + token
  };
  const options = {
    "contentType": "application/json",
    "headers": headers,
    "muteHttpExceptions": true
  };  
  //SELECT句で使われている列名を保存（selects[]：0:テーブル名,1:列名,2:対象テーブルがtables[]の何行目か？後で取得するため初期値0を代入,3:列名に集計関数がある場合にその集計関数）
  let selects = []; 
  let selectM = select;
  let kct = 0; 
  do {
    kct = 0;
    let tmp_sel = "";
    if (selectM.search(',') != -1) {
      tmp_sel = selectM.substring(0, selectM.search(','));
      selectM = selectM.substring(selectM.search(',')+1);
      kct = 1;
    } else {
      tmp_sel = selectM;
    }    
    const gby = ['COUNT','SUM','MAX','MIN','AVG'];
    let grp = "";
    for (const hby of gby) {
      let reg = new RegExp(hby, "i");
      if (tmp_sel.search(reg) != -1) {
        tmp_sel = tmp_sel.substring(tmp_sel.search(reg)+hby.length+1, tmp_sel.search(/\)/));
        grp = hby;
      }
    }
    if (tmp_sel.search(/\./) != -1) {
      let arr = tmp_sel.split(".");
      selects.push([arr[0],arr[1],0,grp]);
    } else {
      if (tmp_sel.search(/\*/) != -1) {
        selects.push(['*','*',0,grp]);
      } else {
        selects.push([tables[0][0],tmp_sel,0,grp]);
      }
    }
  } while(kct);
  selects = [...new Set(selects.map(JSON.stringify))].map(JSON.parse); //重複削除

  //JOIN条件を式単位に分解してそこで使われている列名を保存（joijois[]：0:テーブル名,1:列名,2:対象テーブルがtables[]の何行目か？後で取得するため初期値0を代入）
  let joijois = [];
  for (let ct = 0; ct < tables.length; ct++) {
    let tmp_on1 = tables[ct][3].trim();
    if (tmp_on1 != "") {
      tmp_on1 = tmp_on1.replaceAll('\(','').replaceAll('\)','');
      let jct = 1;
      while(jct) {
        let tmp_on2="";
        if (tmp_on1.search(/ AND /i) != -1) {
          tmp_on2=tmp_on1.substring(0, tmp_on1.search(/ AND /i)).trim();
          tmp_on1=tmp_on1.substring(tmp_on1.search(/ AND /i)+5).trim();
        } else if (tmp_on1.search(/ OR /i) != -1) {
          tmp_on2=tmp_on1.substring(0, tmp_on1.search(/ OR /i)).trim();
          tmp_on1=tmp_on1.substring(tmp_on1.search(/ OR /i)+4).trim();
        } else {
          tmp_on2=tmp_on1.trim();
          jct=0;
        }
        let tmp1=""; //左辺
        let tmp2=""; //右辺
        let equ = ['!=','>=','<=','<>','>','<','='];
        for (const equ4 of equ) {
          let reg4 = new RegExp(equ4, "i");
          if (tmp_on2.search(reg4) != -1) {
            tmp1 = tmp_on2.substring(0, tmp_on2.search(reg4)).trim();
            tmp2 = tmp_on2.substring(tmp_on2.search(reg4)+equ4.length).trim();
            break;
          }
        }
        if (tmp1.search(/\./) != -1) {
          let arr = tmp1.split(".");
          joijois.push([arr[0],arr[1],0]);
        } else {
          joijois.push([tables[0][0],tmp1,0]);
        }
        if (tmp2.search(/\./) != -1) {
          let arr = tmp2.split(".");
          joijois.push([arr[0],arr[1],0]);          
        } else {
          joijois.push([tables[0][0],tmp2,0]);
        }
      }
      tables[ct][3] = tables[ct][3].replaceAll('=','==').replaceAll('!==','!=').replaceAll('>==','>=').replaceAll('<==','<=').replaceAll(/ AND /ig,' && ').replaceAll(/ OR /ig,' || ');
    }
  }
  joijois = [...new Set(joijois.map(JSON.stringify))].map(JSON.parse); //重複削除

  //WHERE句を式単位に分解してそこで使われている列名を保存（wheres[]：0:テーブル名,1:列名）
  let wheres = [];
  let where1=where.replaceAll('\(','').replaceAll('\)','');
  let xct=0;
  if (where1 != "") { xct=1; }
  while(xct) {
    let where2="";
    if (where1.search(/ AND /i) != -1 || where1.search(/ OR /i) != -1) {
      if (where1.search(/ AND /i) == -1) {
        where2=where1.substring(0, where1.search(/ OR /i)).trim();
        where1=where1.substring(where1.search(/ OR /i)+4).trim();
      } else if (where1.search(/ OR /i) == -1) {
        where2=where1.substring(0, where1.search(/ AND /i)).trim();
        where1=where1.substring(where1.search(/ AND /i)+5).trim();
      } else if (where1.search(/ AND /i) > where1.search(/ OR /i)) {
        where2=where1.substring(0, where1.search(/ OR /i)).trim();
        where1=where1.substring(where1.search(/ OR /i)+4).trim();
      } else {
        where2=where1.substring(0, where1.search(/ AND /i)).trim();
        where1=where1.substring(where1.search(/ AND /i)+5).trim();
      }
    } else {
      where2=where1.trim();
      xct=0;
    }
    let tmp1=""; //左辺
    let tmp2=""; //右辺
    let equ = ['!=','>=','<=','<>','>','<','=','LIKE','IS'];
    for (const equ4 of equ) {      
      let reg4 = new RegExp(equ4, "i");
      if (where2.search(reg4) != -1) {
        tmp1 = where2.substring(0, where2.search(reg4)).trim();
        tmp2 = where2.substring(where2.search(reg4)+equ4.length).trim();
        break;
      }
    }
    if (tmp1.search(/\./) != -1) {
      let arr = tmp1.split(".");
      wheres.push([arr[0],arr[1]]);
    } else {
      wheres.push([tables[0][0],tmp1]);
    }
    if (tmp2.search(/\./) != -1) {
      let arr = tmp2.split(".");
      wheres.push([arr[0],arr[1]]);
    }
  }
  where = where.replaceAll('=','==').replaceAll('!==','!=').replaceAll('>==','>=').replaceAll('<==','<=').replaceAll(/ AND /ig,' && ').replaceAll(/ OR /ig,' || ');
  where = where.replaceAll(/IS {1,}NULL/ig,"== ''").replaceAll(/IS {1,}NOT {1,}NULL/ig,"!= ''");
  wheres = [...new Set(wheres.map(JSON.stringify))].map(JSON.parse); //重複削除

  //Query作成
  for (let ct = 0; ct < tables.length; ct++) {
    //SELECT
    let fields = tables[ct][1].split(",");
    let tstr = "";
    let istr = "";
    for (const sele of selects) {
      if ( (sele[0] == '*' || sele[0] == tables[ct][0]) && (sele[3] != "COUNT") ) {
        sele[2] = ct;
        for (let ft = 0; ft < fields.length; ft++) {
          if (sele[1] == fields[ft]) {
            tstr += (","+sele[1]);
            istr += (","+ConvertToLetter(ft));
            break;
          } else if (sele[1] == "*") {
            tstr += (","+fields[ft]);
            istr += (","+ConvertToLetter(ft));
          }
        }
      }
    }
    tables[ct][5] += tstr.substring(1); //見出し
    tables[ct][6] += istr.substring(1); //Query関数
    //{SELECTの列名}に{JOIN条件式で必要とする列名joijois[]}が含まれていない場合は{JOIN条件式で必要とする列名}を{SELECTの列名}に追加する
    for (const jois of joijois) {
      if (jois[0] == tables[ct][0]) {
        jois[2] = ct;
        let t_hoge = tables[ct][5].split(",");
        if (t_hoge.indexOf(jois[1]) == -1) {
          if (fields.indexOf(jois[1]) == -1) {
            throw new Error('（'+jois[0]+m29+jois[1]+m30+'）\n');
          } 
          if (tables[ct][6] == 'SELECT ') {
            tables[ct][5] += (jois[1]);
            tables[ct][6] += (ConvertToLetter(fields.indexOf(jois[1])));
          } else {
            tables[ct][5] += (","+jois[1]);
            tables[ct][6] += (","+ConvertToLetter(fields.indexOf(jois[1])));
          }
        }
      }
    }
    //{SELECTの列名}に{WHEREで必要とする列名wheres[]}が含まれていない場合は{WHEREで必要とする列名}を{SELECTの列名}に追加する
    for (const wher of wheres) {
      if (wher[0] == tables[ct][0]) {
        let t_hoge = tables[ct][5].split(",");
        if (t_hoge.indexOf(wher[1]) == -1) {
          if (fields.indexOf(wher[1]) == -1) {
            throw new Error('（'+wher[0]+m29+wher[1]+m30+'）\n');
          }
          if (tables[ct][6] == 'SELECT ') {
            tables[ct][5] += (wher[1]);
            tables[ct][6] += (ConvertToLetter(fields.indexOf(wher[1])));
          } else {
            tables[ct][5] += (","+wher[1]);
            tables[ct][6] += (","+ConvertToLetter(fields.indexOf(wher[1])));
          }
        }
      }
    }
  }

  let arrays = [];
  for (let i = 0; i < tables.length; i++) {
    //Google Visualization API
    let URL = BASE_URL + SS_ID + "/gviz/tq?gid=" + tables[i][7] + "&range=" + tables[i][4] + "&tqx=out:json&tq=";
    let response = UrlFetchApp.fetch(URL + encodeURIComponent(tables[i][6]), options); //テーブル結合がある場合のクエリ関数はSELECTのみ送信する。SELECT以外のWHERE等のクエリは後で独自処理する。
    let jobj = get_jobj(response);
    let tmp_array = [];
    tmp_array = tmp_array.concat(get_array2(jobj,tables[i][0],tables[i][5]));
    arrays[i] = tmp_array;
  }
  //JOINの独自処理
  let outarray = JSON.parse(JSON.stringify(arrays[0]));
  for (let i = 0; i < tables.length-1; i++) {
    let tmparray = [];
    let strkk = tables[i+1][3];
    for (const jojo of joijois) {
      if (jojo[2] <= i) {
        const ren1 = jojo[0]+"."+jojo[1];
        const ren2 = "helloworld['"+ren1+"']";
        strkk = strkk.replaceAll(ren1,ren2);
      }
    }
    for (const jojo of joijois) {
      if (jojo[2] == (i+1)) {
        const ren1 = jojo[0]+"."+jojo[1];
        const ren2 = "hellowork['"+ren1+"']";
        strkk = strkk.replaceAll(ren1,ren2);
      }
    }
    if (tables[i][2] == 'LEFT') {
      const fields2 = tables[i+1][5].split(",");
      let values2 = {};
      for (let field2 of fields2) {
        const ren2 = tables[i+1][0]+"."+field2;
        values2[ren2]="";
      }
      for (let key1 in outarray) {
        let ck=0;
        for (let key2 in arrays[i+1]) {
          if (hellotheworld(outarray[key1],arrays[i+1][key2],strkk)) {
            tmparray.push({ ...outarray[key1], ...arrays[i+1][key2] });
            ck++;
          }
        }
        if (ck == 0) {
          tmparray.push({ ...outarray[key1], ...values2});
        }
      }
      outarray = JSON.parse(JSON.stringify(tmparray));
    } else if (tables[i][2] == 'RIGHT') {
      const fields3 = tables[i][5].split(",");
      let values3 = {};
      for (let field3 of fields3) {
        const ren3 = tables[i][0]+"."+field3;
        values3[ren3]="";
      }
      for (let key2 in arrays[i+1]) {
        let ck=0;
        for (let key1 in outarray) {
          if (hellotheworld(outarray[key1],arrays[i+1][key2],strkk)) {
            tmparray.push({ ...outarray[key1], ...arrays[i+1][key2] });
            ck++;
          }
        }
        if (ck == 0) {
          tmparray.push({ ...values3, ...arrays[i+1][key2]});
        }
      }
      outarray = JSON.parse(JSON.stringify(tmparray));
    } else if (tables[i][2] == 'INNER' || tables[i][2] == 'JOIN') {
      for (let key1 in outarray) {
        for (let key2 in arrays[i+1]) {
          if (hellotheworld(outarray[key1],arrays[i+1][key2],strkk)) {
            tmparray.push({ ...outarray[key1], ...arrays[i+1][key2] });
          }
        }
      }
      outarray = JSON.parse(JSON.stringify(tmparray));
    } else if (tables[i][2] == 'FULL') {
      const fields2 = tables[i+1][5].split(",");
      let values2 = {};
      for (let field2 of fields2) {
        const ren2 = tables[i+1][0]+"."+field2;
        values2[ren2]="";
      }
      const fields3 = tables[i][5].split(",");
      let values3 = {};
      for (let field3 of fields3) {
        const ren3 = tables[i][0]+"."+field3;
        values3[ren3]="";
      }
      for (let key1 in outarray) {
        let ck=0;
        for (let key2 in arrays[i+1]) {
          if (hellotheworld(outarray[key1],arrays[i+1][key2],strkk)) {
            tmparray.push({ ...outarray[key1], ...arrays[i+1][key2] });
            ck++;
          }
        }
        if (ck == 0) {
          tmparray.push({ ...outarray[key1], ...values2});
        }
      }
      for (let key2 in arrays[i+1]) {
        let ck=0;
        for (let key1 in outarray) {
          if (hellotheworld(outarray[key1],arrays[i+1][key2],strkk)) {
            ck++;
          }
        }
        if (ck == 0) {
          tmparray.push({ ...values3, ...arrays[i+1][key2]});
        }        
      }
      outarray = JSON.parse(JSON.stringify(tmparray));
    } else if (tables[i][2] == 'CROSS' || tables[i][2] == ',') {
      for (let key1 in outarray) {
        for (let key2 in arrays[i+1]) {
          tmparray.push({ ...outarray[key1], ...arrays[i+1][key2] });
        }
      }
      outarray = JSON.parse(JSON.stringify(tmparray));
    }
  }

  //WHERE句の独自処理
  if (where) {
    //WHERE句のLIKE対応
    where = like_to_script(where);
    for (const heres of wheres) {
      const hogehoge = heres[0]+"."+heres[1];
      const gehogeho = "wiz['"+hogehoge+"']";
      where = where.replaceAll(hogehoge,gehogeho);
    }
    const areare = outarray.filter(function(wiz){
      return eval(where);
    });
    outarray = JSON.parse(JSON.stringify(areare));
  }

  //GROUP BY句（HAVING句）の独自処理
  if (group) { outarray = mygrhaving(outarray,select,group,having); }

  //ORDER BY句の独自処理
  if (order) {
    order = order.substring(order.search(/by/i)+2);
    const orders = [];
    let kct = 0;
    do {
      kct = 0;
      let tmp_or="";
      if (order.search(',') != -1) {
        tmp_or = order.substring(0, order.search(',')).trim();
        order = order.substring(order.search(',')+1).trim();
        kct = 1;
      } else {
        tmp_or = order;
      }
      if (tmp_or.search(/DESC/i) != -1) {
        orders.push([tmp_or.substring(0,tmp_or.search(/DESC/i)).trim(),-1]);
      } else if (tmp_or.search(/ASC/i) != -1) {
        orders.push([tmp_or.substring(0,tmp_or.search(/ASC/i)).trim(),1]);
      } else {
        orders.push([tmp_or.trim(),1]);
      }
    } while(kct);
    outarray.sort((a,b) => {
      for (const ddder of orders) {
        if (a[ddder[0]] < b[ddder[0]]) return (-1*ddder[1]);
        if (a[ddder[0]] > b[ddder[0]]) return (1*ddder[1]);
      }
      return 0;
    });
  }

  //見出し（ラベル）作成
  let nok = "";
  for (const outs of selects) {
    if (outs[0]=="*" && !outs[3]) {
      for (const key in tables) {
        const tmp_cols = tables[key][1].split(",");
        for (const t_key in tmp_cols) {
          const lov = tables[key][0]+"."+tmp_cols[t_key];
          nok += (","+lov);
        }
      }
      break;
    }
    if (outs[1]=="*" && !outs[3]) {
      const tmp_cols = tables[outs[2]][1].split(",");
      for (const t_key in tmp_cols) {
        const lov = outs[0]+"."+tmp_cols[t_key];
        nok += (","+lov);
      }
    } else if (outs[3]=="COUNT") {
      nok += ",COUNT(*)";
    } else {
      const lab = outs[0]+"."+outs[1];
      nok += (","+lab);
    }
  }
  let areareee = nok.substring(1).split(",");

  //SELECT通りにカラムを並び替え
  let final_array = [];
  for (const outa of outarray) {
    let semi_array =[];
    for (const arare of areareee) {
      if (outa[arare] == null) {
        semi_array.push('');
      } else {
        semi_array.push(outa[arare]);
      }
    }
    final_array.push(semi_array);
  }
  //テーブル名.列名からテーブル名を除外して列名だけにする
  nok = ","+areareee.join();
  nok = nok.replaceAll(",",", ");
  nok = nok.replaceAll("."," .");
  nok = nok.replaceAll(/ \S{1,} \./g,"");
  final_array.unshift(nok.substring(1).split(","));

  return final_array;
}
//------------------------------------------------------------------------------------------------------------
//Google Visualization API のレスポンス
function get_jobj(response) {
  let data = response.getContentText();
  data = data.split("google.visualization.Query.setResponse(")[1];
  data = data.substr(0, data.length - 2);
  let jobj = JSON.parse(data);
  return jobj;
}
//------------------------------------------------------------------------------------------------------------
//Google Visualization API のレスポンス（JSON）を[配列]に変換
function get_array(jobj) {
  let rowlen = jobj["table"]["rows"].length;
  let collen = jobj["table"]["cols"].length;
  let array = [];
  for (let i = 0; i < rowlen; i++) {
    var values = [];
    for (let j = 0; j < collen; j++) {
      if (jobj["table"]["rows"][i]["c"][j] == null) {
        values.push('');
      } else {
        if (jobj["table"]["rows"][i]["c"][j]["f"] == null) {
          values.push(jobj["table"]["rows"][i]["c"][j]["v"]);
        } else {
          values.push(jobj["table"]["rows"][i]["c"][j]["f"]);
        }
      }
    }
    array.push(values);
  }
  return array;
}
//------------------------------------------------------------------------------------------------------------
//Google Visualization API のレスポンス（JSON）を{連想配列}に変換
function get_array2(jobj,t_name,t_cals) {
  let sCals = t_cals.split(",");
  let rowlen = jobj["table"]["rows"].length;
  let collen = jobj["table"]["cols"].length;
  let array = [];
  for (let i = 0; i < rowlen; i++) {
    let values = {};
    for (let j = 0; j < collen; j++) {
      let rens = t_name+"."+sCals[j];
      if (jobj["table"]["rows"][i]["c"][j] == null) {
        values[rens] = '';
      } else {
        if (jobj["table"]["rows"][i]["c"][j]["f"] == null) {
          if (values[rens] = jobj["table"]["rows"][i]["c"][j]["v"] == null) {
            values[rens] = '';
          } else {
            values[rens] = jobj["table"]["rows"][i]["c"][j]["v"];
          }
        } else {
          values[rens] = jobj["table"]["rows"][i]["c"][j]["f"];
        }
      }
    }
    array.push(values);
  }
  return array;
}
//------------------------------------------------------------------------------------------------------------
//カラムをアルファベットから数（カラム数：1～）に変換（例：A→1,B→2,C→3,...,Z→26,AA→27）
function AbcToNumber(strColumn) {
  var iNum = 0;
  var tmp = 0;  
  strColumn = strColumn.toUpperCase();
  for (i = strColumn.length - 1; i >= 0; i--) {
    tmp = strColumn.charCodeAt(i) - 65;
    if(i != strColumn.length - 1) {
      tmp = (tmp + 1) * Math.pow(26,(i + 1));
    }
    iNum = iNum + tmp;
  }
  return iNum + 1;
}
//------------------------------------------------------------------------------------------------------------
//カラムを数（配列の添字：0～）からアルファベットに変換（例：0→A,1→B,2→C,...,25→Z,26→AA）
function ConvertToLetter(iCol) {
  let str = "";
  let iAlp = 0;
  let iRem = 0;
  iAlp = parseInt((iCol / 26), 10);
  iRem = iCol - (iAlp * 26);
  if (iAlp > 0) {
    str = String.fromCharCode(iAlp + 64);
  }
  if (iRem >= 0) {
    str = str + String.fromCharCode(iRem + 65);
  }
  return str;
}
//------------------------------------------------------------------------------------------------------------
//TableA(helloworld) [FULL|LEFT|RIGHT] [OUTER|INNER] JOIN TableB(hellowork) ON (条件:strhogehoge)
function hellotheworld(helloworld, hellowork, strhogehoge){
  const strgg = "if ("+strhogehoge+") { true; } else { false; }";
  return eval(strgg);
}
//------------------------------------------------------------------------------------------------------------
//WHERE句のBETWEENとINの対応
function sql_to_script(where) {
  //WHERE(NOT BETWEEN対応：X not between A and B → X < A and X > B)
  while (where.search(/NOT {1,}BETWEEN/i) != -1) {
    let nhoge = where.substring(where.search(/\S{1,} {1,}NOT {1,}BETWEEN {1,}/i),where.search(/ {1,}NOT {1,}BETWEEN {1,}/i));
    let ntmp_where = where.substring(where.search(/BETWEEN/ig)+7);
    let ntmp_e1 = ntmp_where.substring(0,ntmp_where.search(/ AND /ig)).trim();
    ntmp_where = ntmp_where.substring(ntmp_where.search(/ AND /ig)+5).trim();
    let ntmp_e2 = ntmp_where.substring(0,ntmp_where.search(' '));
    let ntmp_q1 = " ( "+nhoge+" < "+ntmp_e1+" AND "+nhoge+" > "+ntmp_e2+" ) ";
    let ntmp_q2 = nhoge+" {1,}NOT {1,}BETWEEN {1,}"+ntmp_e1+" {1,}AND {1,}"+ntmp_e2;
    let nreg = new RegExp(ntmp_q2,"i");
    where = where.replace(nreg,ntmp_q1);
  }
  //WHERE(BETWEEN対応：X between A and B → X >= A and X <= B)
  while (where.search(/BETWEEN/i) != -1) {
    let hoge = where.substring(where.search(/\S{1,} {1,}BETWEEN {1,}/i),where.search(/ {1,}BETWEEN {1,}/i));
    let tmp_where = where.substring(where.search(/BETWEEN/ig)+7);
    let tmp_e1 = tmp_where.substring(0,tmp_where.search(/ AND /ig)).trim();
    tmp_where = tmp_where.substring(tmp_where.search(/ AND /ig)+5).trim();
    let tmp_e2 = tmp_where.substring(0,tmp_where.search(' '));    
    let tmp_q1 = " ( "+hoge+" >= "+tmp_e1+" AND "+hoge+" <= "+tmp_e2+" ) ";
    let tmp_q2 = hoge+" {1,}BETWEEN {1,}"+tmp_e1+" {1,}AND {1,}"+tmp_e2;
    let reg = new RegExp(tmp_q2,"i");
    where = where.replace(reg,tmp_q1);
  }
  //WHERE(NOT IN対応：X not in (A,B,..) → X != A and X != B and ...)
  while (where.search(/NOT {1,}IN/i) != -1) {
    let iii = where.search(/\S{1,} {1,}NOT {1,}IN {1,}/i);
    let incol = where.substring(iii,where.search(/ {1,}NOT {1,}IN {1,}/i));
    let tmp_in = where.substring(iii);
    let jjj = tmp_in.search(/\)/);
    tmp_in = tmp_in.substring(tmp_in.search(/\(/)+1,jjj);
    let tmptmp = "";
    let in_tmp = tmp_in.split(",");
    for (const inin of in_tmp) {
      tmptmp = tmptmp + incol+" != "+inin+" AND ";
    }
    tmptmp = " ( "+tmptmp.substring(0,tmptmp.length-4)+" ) ";
    where = where.replace(where.substring(iii,iii+jjj+1),tmptmp);
  }
  //WHERE(IN対応：X in (A,B,...) → X = A or X = B or ...)
  while (where.search(/ {1,}IN/i) != -1) {
    let iii = where.search(/\S{1,} {1,}IN {1,}/i);
    let incol = where.substring(iii,where.search(/ {1,}IN {1,}/i));
    let tmp_in = where.substring(iii);
    let jjj = tmp_in.search(/\)/);
    tmp_in = tmp_in.substring(tmp_in.search(/\(/)+1,jjj);
    let tmptmp = "";
    let in_tmp = tmp_in.split(",");
    for (const inin of in_tmp) {
      tmptmp = tmptmp + incol+" = "+inin+" OR ";
    }
    tmptmp = " ( "+tmptmp.substring(0,tmptmp.length-3)+" ) ";
    where = where.replace(where.substring(iii,iii+jjj+1),tmptmp);
  }
  return where;
}
//------------------------------------------------------------------------------------------------------------
//WHERE句のLIKE対応
function like_to_script(where) {
  //X LIKE '%A%' （前方後方一致）
  while(where.search(/ {1,}LIKE {1,}'%[^%']+%'/i) != -1) {
    let hoge = Math.random().toString(32).substring(2);
    while (where.search(hoge) != -1) {
      hoge = Math.random().toString(32).substring(2);
    }
    where = where.replace(/ {1,}LIKE {1,}'%/i,hoge);
    where = where.substring(0,where.search(hoge)+hoge.length) + where.substring(where.search(hoge)+hoge.length).replace("%'","'\) != -1 ");
    where = where.replace(hoge,".search\('");
  }
  //X LIKE '%A' （後方一致）
  while(where.search(/ {1,}LIKE {1,}'%/i) != -1) {
    let hoge = Math.random().toString(32).substring(2);
    while (where.search(hoge) != -1) {
      hoge = Math.random().toString(32).substring(2);
    }
    where = where.replace(/ {1,}LIKE {1,}'%/i,hoge);
    where = where.substring(0,where.search(hoge)+hoge.length) + where.substring(where.search(hoge)+hoge.length).replace("'","'\)");
    where = where.replace(hoge,".endsWith\('");
  }
  //X LIKE 'A%' （前方一致）
  while(where.search(/ {1,}LIKE {1,}'/i) != -1) {
    let hoge = Math.random().toString(32).substring(2);
    while (where.search(hoge) != -1) {
      hoge = Math.random().toString(32).substring(2);
    }
    where = where.replace(/ {1,}LIKE {1,}'/i,hoge);
    where = where.substring(0,where.search(hoge)+hoge.length) + where.substring(where.search(hoge)+hoge.length).replace("%'","'\)");
    where = where.replace(hoge,".startsWith\('");
  }
  return where;
}
//------------------------------------------------------------------------------------------------------------
//GROUP BY句（HAVING句）対応
function mygrhaving(t_array,select,group,having) {
  Object.prototype.values = function(){ var o=this; var r=[]; for(var k in o) if(o.hasOwnProperty(k)) { r.push(o[k]) } return r; };
  Array.prototype.groupBy = function(keys,nomKeys,cotKeys,sumKeys,avgKeys,maxKeys,minKeys) {
    var hash = this.reduce(function(res,data) {
      var key = keys.reduce(function(s,k) {
        s += data[k];
        return s;
      },'');
      if(!(key in res)) {
        var keyList = keys.reduce(function(h,k) { h[k] = data[k]; return h; },{});
        res[key] = cotKeys.reduce(function(h,k) { h[k] = 0; return h; }, keyList);
        res[key] = sumKeys.reduce(function(h,k) { h[k] = 0; return h; }, keyList);
        res[key] = avgKeys.reduce(function(h,k) { h[k] = 0; return h; }, keyList);
        res[key] = maxKeys.reduce(function(h,k) { h[k] = data[k]; return h; }, keyList);
        res[key] = minKeys.reduce(function(h,k) { h[k] = data[k]; return h; }, keyList);
      }
      nomKeys.forEach(function(k){ res[key][k] = data[k]; });
      cotKeys.forEach(function(k){ res[key][k] += 1; });
      sumKeys.forEach(function(k){ res[key][k] += Number.parseInt(data[k]); });
      avgKeys.forEach(function(k){ res[key][k] += Number.parseInt(data[k]); });
      maxKeys.forEach(function(k){ res[key][k] = Math.max(res[key][k], data[k]); } );
      minKeys.forEach(function(k){ res[key][k] = Math.min(res[key][k], data[k]); } );
      return res;
    },{});
    return hash.values();
  }

  let groups = group.substring(group.search(/BY/i)+2).trim().split(",");
  let groupst = "";
  for (const element of groups) {
    groupst = groupst + "," + element.trim();
  }
  if (groupst) { groupst = groupst.substring(1); }
  let selectst = select.split(",");
  let sel = { "SEL": "", "COUNT": "", "SUM": "", "MAX": "", "MIN": "", "AVG": "" };
  for (const element of selectst) {
    const gby = ['COUNT','SUM','MAX','MIN','AVG'];
    let cst = true;
    for (const hby of gby) {
      let reg = new RegExp(hby, "i");
      if (element.search(reg) != -1) {
        sel[hby] = sel[hby] + "," + element.substring(element.search(reg)+hby.length+1, element.search(/\)/)).trim();
        cst = false; 
      }
    }
    if (cst) { sel['SEL'] = sel['SEL'] + "," + element.trim(); }
  }
  t_array = t_array.groupBy(groupst.split(","),groupst.split(","),['COUNT\(\*\)'],sel["SUM"].substring(1).split(","),sel["AVG"].substring(1).split(","),sel["MAX"].substring(1).split(","),sel["MIN"].substring(1).split(","));
  if (sel["AVG"]) {
    for (const t_arra of t_array) {
      for (const avgct of sel["AVG"].substring(1).split(",")) { t_arra[avgct] = Math.round((t_arra[avgct] / t_arra['COUNT\(\*\)'])*100)/100; }
    }
  }
  //HAVING句
  if (having) {
    //having = having.replaceAll('\(','').replaceAll('\)','');
    having = sql_to_script(having);
    having = having.replaceAll('=','==').replaceAll('!==','!=').replaceAll('>==','>=').replaceAll('<==','<=').replaceAll(/ AND /ig,' && ').replaceAll(/ OR /ig,' || ');
    having = having.replaceAll(/IS {1,}NULL/ig,"== ''").replaceAll(/IS {1,}NOT {1,}NULL/ig,"!= ''");
    having = like_to_script(having);
    let set = new Set();
    let has=having;
    let has2;
    do {
      if (has.indexOf("&&") != -1 || has.indexOf("||") != -1) {
        if (has.indexOf("&&") == -1) {
          has2 = has.substring(0, has.indexOf("||")).trim();
          has = has.substring(has.indexOf("||")+2).trim();
        } else if (has.indexOf("||") == -1) {
          has2 = has.substring(0, has.indexOf("&&")).trim();
          has = has.substring(has.indexOf("&&")+2).trim();
        } else if (has.indexOf("&&") > has.indexOf("||")) {
          has2 = has.substring(0, has.indexOf("||")).trim();
          has = has.substring(has.indexOf("||")+2).trim();
        } else {
          has2 = has.substring(0, has.indexOf("&&")).trim();
          has = has.substring(has.indexOf("&&")+2).trim();
        }
      } else { has2 = has; }
      let tmp1=has2;
      let tmp2="";
      let equ = ['!=','>=','<=','<>','>','<','=='];
      for (const equ4 of equ) {
        let reg4 = new RegExp(equ4, "i");
        if (has2.search(reg4) != -1) {
          tmp1 = has2.substring(0, has2.search(reg4)).trim();
          tmp2 = has2.substring(has2.search(reg4)+equ4.length).trim();
          break;
        }
      }
      const gby = ['SUM','MAX','MIN','AVG'];
      let tmp1_f = true;
      let tmp2_f = true;
      for (const hby of gby) {
        let reg = new RegExp(hby, "i");
        if (tmp1.search(reg) != -1) {
          if (!set.has(tmp1)) {
            having = having.replaceAll(tmp1,"wiz['"+tmp1.substring(tmp1.search(reg)+hby.length+1, tmp1.search(/\)/)).trim()+"']");
            set.add(tmp1);
          }
          tmp1_f = false;
        }
        if (tmp2.search(reg) != -1) {
          if (!set.has(tmp2)) {
            having = having.replaceAll(tmp2,"wiz['"+tmp2.substring(tmp2.search(reg)+hby.length+1, tmp2.search(/\)/)).trim()+"']");
            set.add(tmp2);
          }
          tmp2_f = false;
        }
      }
      if (tmp1.search(/COUNT\(/i) != -1) {
        if (!set.has(tmp1)) {
          having = having.replaceAll(tmp1,"wiz['COUNT(*)']");
          set.add(tmp1);
        }
        tmp1_f = false;
      }
      if (tmp2.search(/COUNT\(/i) != -1) {
        if (!set.has(tmp2)) {
          having = having.replaceAll(tmp2,"wiz['COUNT(*)']");
          set.add(tmp2);
        }
        tmp2_f = false;
      }
      if (tmp1_f) { 
        if (!set.has(tmp1)) {
          having = having.replaceAll(tmp1,"wiz['"+tmp1+"']");
          set.add(tmp1);
        }
      }
      if (tmp2.search(/\./) != -1 && tmp2_f) {
        if (!set.has(tmp2)) {
          having = having.replaceAll(tmp2,"wiz['"+tmp2+"']");
          set.add(tmp2);
        }
      }
    } while (has.indexOf("&&") != -1 || has.indexOf("||") != -1);
    t_array = t_array.filter(function(wiz){ return eval(having); });
  }

  return t_array;
}
//------------------------------------------------------------------------------------------------------------
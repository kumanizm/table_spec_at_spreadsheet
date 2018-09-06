
var _config = {
  'DB_HOST' : "<MySQLのホスト>",
  'DB_NAME' : "<MySQLのDB名>",
  'DB_USER' : "<MySQLの接続ユーザ>",
  'DB_PASS' : "<MySQLの接続パス>",
  'DB_PORT' : "<MySQLの接続ポート>",
  
  'SS_FILE_ID' : "<スプレッドシートのファイルID>",
  'SS_LIST_NAME' : "テーブル一覧",
};

var SPREADSHEET;
var LIST_SHEET ;
var tables = [];

function get_list_(stmt) {
  tables = [];

  //クエリを記載
  var rs = stmt.executeQuery('show tables from ' + _config['DB_NAME'] + ';');
  while(rs.next()) {
    //getStringで列名を指定して取得
    tables.push(rs.getString("Tables_in_" + _config['DB_NAME']));
  }
  rs.close();
}

function set_list_sheet_(stmt) {

  LIST_SHEET.clear();
  LIST_SHEET.getRange("A1").setValue("No").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  LIST_SHEET.getRange("B1").setValue("テーブル名").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  
  for (var i=0; i<tables.length; i++) {
    var table_ss = SPREADSHEET.getSheetByName(tables[i]);
    LIST_SHEET.getRange("A"+(i+2)).setValue(i+1).setBorder(true, true, true, true, true, true).autoResizeColumn(1);
    LIST_SHEET.getRange("B"+(i+2)).setValue("=HYPERLINK(\""+"#gid="+ table_ss.getSheetId() +"\", \""+tables[i]+"\")").setBorder(true, true, true, true, true, true).autoResizeColumn(2);
  }

}

function set_table_sheet_(TABLE, stmt) {
  var table_ss = SPREADSHEET.getSheetByName(TABLE);
  if (table_ss == null) {
    table_ss = SPREADSHEET.insertSheet(TABLE);
  }
  
  table_ss.clear();
  
  var j=1;
  
  table_ss.getRange("A"+j).setValue("テーブル情報").setFontWeight('bold');
  j++;
  table_ss.getRange("A"+j).setValue("論理名").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("B"+j).setValue("物理名").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("C"+j).setValue("エンジン").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("D"+j).setValue("文字コード").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("E"+j).setValue("作成日").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("F"+j).setValue("更新日").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  
  var str_query = "show table status like '" + TABLE + "'";
  var rs = stmt.executeQuery(str_query);
  k=1;
  while(rs.next()) {
    //getStringで列名を指定して取得
    table_ss.getRange("A"+(j+1)).setValue(rs.getString("Comment")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("B"+(j+1)).setValue(rs.getString("Name")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("C"+(j+1)).setValue(rs.getString("Engine")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("D"+(j+1)).setValue(rs.getString("Collation")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("E"+(j+1)).setValue(rs.getString("Create_time")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("F"+(j+1)).setValue(rs.getString("Update_time")).setBorder(true, true, true, true, true, true);
    
    j++;
    k++;
  }
  rs.close();
  
  j=j+2;
  
  table_ss.getRange("A"+j).setValue("カラム情報").setFontWeight('bold');
  j++;
  table_ss.getRange("A"+j).setValue("No").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("B"+j).setValue("論理名").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("C"+j).setValue("物理名").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("D"+j).setValue("データ型").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("E"+j).setValue("Null").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("F"+j).setValue("Key").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("G"+j).setValue("Extra").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("H"+j).setValue("デフォルト").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("I"+j).setValue("文字コード").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("J"+j).setValue("備考").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  
  
  var str_query = 'show full columns from ' + TABLE;
  var rs = stmt.executeQuery(str_query);
  
  while(rs.next()) {
    //getStringで列名を指定して取得
    table_ss.getRange("A"+(j+1)).setValue(j-1).setBorder(true, true, true, true, true, true);
    table_ss.getRange("B"+(j+1)).setValue(rs.getString("Comment")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("C"+(j+1)).setValue(rs.getString("Field")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("D"+(j+1)).setValue(rs.getString("Type")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("E"+(j+1)).setValue(rs.getString("Null")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("F"+(j+1)).setValue(rs.getString("Key")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("G"+(j+1)).setValue(rs.getString("Extra")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("H"+(j+1)).setValue(rs.getString("Default")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("I"+(j+1)).setValue(rs.getString("Collation")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("J"+(j+1)).setValue(rs.getString("Comment")).setBorder(true, true, true, true, true, true);
    
    j++;
  }
  rs.close();

  j=j+2;
  
  table_ss.getRange("A"+j).setValue("インデックス情報").setFontWeight('bold');
  j++;
  table_ss.getRange("A"+j).setValue("No").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("B"+j).setValue("インデックス名").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("C"+j).setValue("カラムリスト").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("D"+j).setValue("主キー").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("E"+j).setValue("ユニーク").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("F"+j).setValue("備考").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  
  var str_query = 'show index from ' + TABLE;
  var rs = stmt.executeQuery(str_query);
  k=1;
  while(rs.next()) {
    //getStringで列名を指定して取得
    table_ss.getRange("A"+(j+1)).setValue(k).setBorder(true, true, true, true, true, true);
    table_ss.getRange("B"+(j+1)).setValue(rs.getString("Key_name")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("C"+(j+1)).setValue(rs.getString("Column_name")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("D"+(j+1)).setValue( rs.getString("Key_name")=="PRIMARY"?"YES":"" ).setBorder(true, true, true, true, true, true);
    table_ss.getRange("E"+(j+1)).setValue(rs.getString("Non_unique")=="0"?"YES":"").setBorder(true, true, true, true, true, true);
    table_ss.getRange("F"+(j+1)).setValue(rs.getString("Comment")).setBorder(true, true, true, true, true, true);
    
    j++;
    k++;
  }
  rs.close();
  j=j+2;
  
  table_ss.getRange("A"+j).setValue("外部キー情報").setFontWeight('bold');
  j++;
  table_ss.getRange("A"+j).setValue("No").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("B"+j).setValue("外部キー名").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("C"+j).setValue("カラムリスト").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("D"+j).setValue("参照先テーブル名").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  table_ss.getRange("E"+j).setValue("参照先カラムリスト").setBorder(true, true, true, true, true, true).setBackgroundRGB(155, 229, 42).setFontWeight('bold');
  
  var str_query = "SELECT * FROM information_schema.key_column_usage WHERE constraint_schema='" + _config['DB_NAME'] + "' AND referenced_table_name='" + TABLE + "';";
  
  var rs = stmt.executeQuery(str_query);
  k=1;
  while(rs.next()) {
    //getStringで列名を指定して取得
    table_ss.getRange("A"+(j+1)).setValue(k).setBorder(true, true, true, true, true, true);
    table_ss.getRange("B"+(j+1)).setValue(rs.getString("CONSTRAINT_NAME")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("C"+(j+1)).setValue(rs.getString("COLUMN_NAME")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("D"+(j+1)).setValue(rs.getString("TABLE_NAME")).setBorder(true, true, true, true, true, true);
    table_ss.getRange("E"+(j+1)).setValue(rs.getString("COLUMN_NAME")).setBorder(true, true, true, true, true, true);
    
    j++;
    k++;
  }
  rs.close();
/*
  // 列幅を変更する
  table_ss.autoResizeColumn(1);
  table_ss.autoResizeColumn(2);
  table_ss.autoResizeColumn(3);
  table_ss.autoResizeColumn(4);
  table_ss.autoResizeColumn(5);
  table_ss.autoResizeColumn(6);
  table_ss.autoResizeColumn(7);
  table_ss.autoResizeColumn(8);
  table_ss.autoResizeColumn(9);
  table_ss.autoResizeColumn(10);
  */
  j=j+2;
Logger.log(table_ss.getMaxRows());
  if (table_ss.getMaxRows()==1000) {    
    table_ss.deleteColumns(11, table_ss.getMaxColumns()-11);
    table_ss.deleteRows(j, (table_ss.getMaxRows()-j));
  }
}

function set_table_list(){

  var con_str = 'jdbc:mysql://' + _config['DB_HOST'] + ':' + _config['DB_PORT'] + '/' + _config['DB_NAME'];
  var user_id = _config['DB_USER'];
  var user_pass = _config['DB_PASS'];
  
  SPREADSHEET = SpreadsheetApp.openById(_config['SS_FILE_ID']);
  if (SPREADSHEET == null) {
    return;
  }

  var LIST_SHEET = SPREADSHEET.getSheetByName(_config['SS_LIST_NAME']);
  if (LIST_SHEET == null) {
    LIST_SHEET = SPREADSHEET.insertSheet(_config['SS_LIST_NAME']);
  }
　　// DBに接続
  var conn = Jdbc.getConnection(con_str, user_id, user_pass);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  get_list_(stmt);

  // ---

  for (var i=0; i<tables.length; i++) {
    set_table_sheet_(tables[i],stmt);
  }
  set_list_sheet_(stmt);

  stmt.close();
  conn.close();

}

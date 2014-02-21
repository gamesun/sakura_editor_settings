// daily_memo.js
// Sakura Editor Macro
// 日付メモファイル生成マクロ
//
// 指定フォルダ下に、最終投稿メモファイルを基に、本日分のメモファイルを作成する。
// メモファイル名形式：「yyyymmdd.txt」


//=======================================
// 設定
//=======================================

//メモファイル保存場所
var MEMO_DIR="d:\\sunyt\\";




//=======================================
// メイン処理
//=======================================

//最新メモファイル名取得
var lastest_filename = get_latest_filename();

//新規メモファイル名生成
var today = new Date();
var today_filename = pad(today.getFullYear(), 4) + 
	pad(today.getMonth() + 1, 2) + pad(today.getDate(), 2) +
	".txt";

if(lastest_filename == "") {
	lastest_filename = today_filename;
}


//最新メモファイルを基に新規メモファイルを作成
//Editor.FileOpen(MEMO_DIR + lastest_filename);
Editor.FileSaveAs(MEMO_DIR + today_filename);




//=======================================
// メソッド定義
//=======================================

//最新メモファイル名取得
function get_latest_filename() {
	var fso=new ActiveXObject("Scripting.FileSystemObject");
	var memo_dir = fso.GetFolder(MEMO_DIR);
	var fset = new Enumerator(memo_dir.Files);

	var latest_filename = "";
	var latest_date_str = "";
	for (;!fset.atEnd();fset.moveNext()) {
		var filename = fset.item().Name;
		var matches = filename.match(/^(\d+).txt$/);
		if( matches != null &&
		(latest_filename == "" || latest_date_str < matches[1])) {
			latest_filename = filename;
			latest_date_str = matches[1];
		}
	}
	return latest_filename;
}


//文字列ゼロ埋め
function pad(value, size) {
	var zeros = "";
	for(var i=0; i < size; i++) zeros += "0";

	var str = (zeros + value);
	return str.substr(str.length - size, size);
}

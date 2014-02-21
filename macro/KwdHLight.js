// KwdHLight.js
// Coder: sunyt
// Sakura Editor Macro
// 指定した複数のキーワードを、ハイライトさせる。 
//


//=======================================
// 設定
//=======================================

var DIR="D:\\sunyt\\01_tool\\sakura\\macro\\KwdHLightList.txt";

var ForReading   = 1; //読み込み
var ForWriting   = 2; //書きこみ（上書きモード）
var ForAppending = 8; //書きこみ（追記モード）

//=======================================
// メイン処理
//=======================================

main:{
	
	Editor.MoveHistSet();
	
	var i;
	var tmp;
	var idx;
	var isSel = Editor.IsTextSelected();
	var src = ReadFile(DIR);
	if (1 == isSel)		//選択中
	{
		var sel = Editor.GetSelectedString(0);
		
		for( i = 0; i < sel.length; i++ )	//illeagl char selected judge
		{
			tmp = sel.charAt(i)
			if(!(	( "0" <= tmp && tmp <= "9")||
					( "a" <= tmp && tmp <= "z")||
					( "A" <= tmp && tmp <= "Z")||
					( "_" == tmp )
				))
			{
				break main;
			}
		}
		
		if (-1 == src)	// empty file
		{	
			var newstr = "\\b" + sel + "\\b";		// \b : 単語境界
		}
		else
		{
			idx = src.search(sel);
			
			if (-1 != idx)		//選択したキーワードもあるなら、削除する
			{
				if ( (idx + sel.length + 2) >= src.length )	// input WORD is the lastst one in src.
				{
					// string.substring( from,to )
					// string の from～to - 1 文字目（最初の文字を 0 番とする）の文字列を返します。
					// 負の値を指定すると 0 番目と見なされます。to を省略すると残りのすべてを返します。
					newstr = src.substring( 0, (idx - 3) );		// (idx-3): del the "|\bWORD"
				}
				else	// input WORD is the first one or in middle.
				{
					newstr = src.substring( 0, (idx - 2) );				// sub the front (end to '|')
					tmp    = src.substring( (idx + sel.length + 3) );	// sub the rear (from '\b')
					newstr = newstr + tmp;								// link front & rear
				}
				Editor.SearchNext("\1", 22);
			}
			else
			{
				var newstr = src + "|\\b" + sel + "\\b";	// add new WORD to the rear of old src.
			}
		}
		WriteFile(DIR, newstr);
		//var newstr = ReadFile(DIR);
		Editor.SearchNext(newstr, 22);		// 22:正規表現で検索する
	}
	else if (0 == isSel)	//非選択状態
	{
		
		if (-1 == src)	// empty file
		{	
			// search a word which is never been used, so that highlight no word.
			// (Because I hadn't found the way to turn highlight off.)
			Editor.SearchNext("\1", 22);
		}
		else
		{
			Editor.SearchNext(src, 22);
		}
	}
	
	Editor.MoveHistPrev();
}

//=======================================
// メソッド定義
//=======================================
function ReadFile(file_path) {
	var fso      = new ActiveXObject('Scripting.FileSystemObject');
	var f_open   = fso.opentextfile(file_path, ForReading);
	var read_all;
	if( ! f_open.AtEndOfStream )
	{
		read_all = f_open.ReadAll();
	}
	else	// empty file
	{
		read_all = -1;
	}
	f_open.close();
	fso = null;
	return read_all;
};

function WriteFile(file_path, str) {
	var fso      = new ActiveXObject('Scripting.FileSystemObject');
	var f_open   = fso.opentextfile(file_path, ForWriting);
	f_open.Write(str);
	f_open.close();
	fso = null;
};


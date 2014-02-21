//キーボードマクロのファイル
Editor.MoveHistSet();
Editor.MoveHistSet();

Editor.SelectWord(0);	// 現在位置の単語選択
var sel = Editor.GetSelectedString(0);
//Editor.SearchNext(sel, 22);
Editor.SearchNext(sel, 19);

Editor.MoveHistPrev();
Editor.MoveHistPrev();
Editor.SelectWord(0);

//ExpandParameter( "$F($y,$x)" );
//TraceOut( "$F($y,$x)", 1 );



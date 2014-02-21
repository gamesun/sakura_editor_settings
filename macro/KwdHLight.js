// KwdHLight.js
// Coder: sunyt
// Sakura Editor Macro
// �w�肵�������̃L�[���[�h���A�n�C���C�g������B 
//


//=======================================
// �ݒ�
//=======================================

var DIR="D:\\sunyt\\01_tool\\sakura\\macro\\KwdHLightList.txt";

var ForReading   = 1; //�ǂݍ���
var ForWriting   = 2; //�������݁i�㏑�����[�h�j
var ForAppending = 8; //�������݁i�ǋL���[�h�j

//=======================================
// ���C������
//=======================================

main:{
	
	Editor.MoveHistSet();
	
	var i;
	var tmp;
	var idx;
	var isSel = Editor.IsTextSelected();
	var src = ReadFile(DIR);
	if (1 == isSel)		//�I��
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
			var newstr = "\\b" + sel + "\\b";		// \b : �P�ꋫ�E
		}
		else
		{
			idx = src.search(sel);
			
			if (-1 != idx)		//�I�������L�[���[�h������Ȃ�A�폜����
			{
				if ( (idx + sel.length + 2) >= src.length )	// input WORD is the lastst one in src.
				{
					// string.substring( from,to )
					// string �� from�`to - 1 �����ځi�ŏ��̕����� 0 �ԂƂ���j�̕������Ԃ��܂��B
					// ���̒l���w�肷��� 0 �Ԗڂƌ��Ȃ���܂��Bto ���ȗ�����Ǝc��̂��ׂĂ�Ԃ��܂��B
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
		Editor.SearchNext(newstr, 22);		// 22:���K�\���Ō�������
	}
	else if (0 == isSel)	//��I�����
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
// ���\�b�h��`
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


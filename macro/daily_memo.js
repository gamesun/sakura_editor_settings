// daily_memo.js
// Sakura Editor Macro
// ���t�����t�@�C�������}�N��
//
// �w��t�H���_���ɁA�ŏI���e�����t�@�C������ɁA�{�����̃����t�@�C�����쐬����B
// �����t�@�C�����`���F�uyyyymmdd.txt�v


//=======================================
// �ݒ�
//=======================================

//�����t�@�C���ۑ��ꏊ
var MEMO_DIR="d:\\sunyt\\";




//=======================================
// ���C������
//=======================================

//�ŐV�����t�@�C�����擾
var lastest_filename = get_latest_filename();

//�V�K�����t�@�C��������
var today = new Date();
var today_filename = pad(today.getFullYear(), 4) + 
	pad(today.getMonth() + 1, 2) + pad(today.getDate(), 2) +
	".txt";

if(lastest_filename == "") {
	lastest_filename = today_filename;
}


//�ŐV�����t�@�C������ɐV�K�����t�@�C�����쐬
//Editor.FileOpen(MEMO_DIR + lastest_filename);
Editor.FileSaveAs(MEMO_DIR + today_filename);




//=======================================
// ���\�b�h��`
//=======================================

//�ŐV�����t�@�C�����擾
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


//������[������
function pad(value, size) {
	var zeros = "";
	for(var i=0; i < size; i++) zeros += "0";

	var str = (zeros + value);
	return str.substr(str.length - size, size);
}

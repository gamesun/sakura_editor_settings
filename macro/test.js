

//=======================================
// �ݒ�
//=======================================

var DIR1="D:\\Zeus.exe";
//var DIR2="D:\\Zeus_xue.exe";
var DIROUTPUT="D:\\tmp.txt";

var ForReading   = 1; //�ǂݍ���
var ForWriting   = 2; //�������݁i�㏑�����[�h�j
var ForAppending = 8; //�������݁i�ǋL���[�h�j

//=======================================
// ���C������
//=======================================

main:{
	var tmp;
	var i;
	
	var src1 = ReadFile(DIR1);
//	var src2 = ReadFile(DIR2);
	
	WriteFile(DIROUTPUT,src1);
//	i = 0;
//	tmp = src1.charAt(i);
	
//	while(tmp){
//		WriteFile(DIROUTPUT,tmp);
//		i++;
//		tmp = src1.charAt(i);
//	}
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
	var f_open   = fso.opentextfile(file_path, ForAppending);
	f_open.Write(str);
	f_open.close();
	fso = null;
};
'
' Sakura File Manager 1.0 beta 1
'
' (c)2005 �Ձ[��
' webmaster@muumoo.jp
' http://muumoo.jp/
'

'
' �ݒ�
'
' �^�C�g��
Dim title
title = "Sakura File Manager"
' �o�[�W����
Dim version
version = "1.0 beta 1"
' ��؂��
Dim separator
separator = "----------------"
' �I��̈�
Dim selection
selection = GetSelection()
' ���łɃt�@�C���}�l�[�W���ɓ����Ă��邩�ۂ�
Dim isFileManagerMode
isFileManagerMode = IsFileManager()

' ���s
Main()

'
' ���C��
'
Sub Main
	
	' �R�}���h���ǂ�������
	If isFileManagerMode Then
		Dim commandLine
		If selection <> "" Then
			commandLine = selection
		Else
			commandLine = GetLine()
		End If
		If Left(commandLine, 1) = "!" Then
			DoCommand(commandLine)
			Exit Sub
		End If
	End If
	
	' �Ώۃp�X���擾���܂��B
	Dim path
	path = GetPath()
	
	' �p�X�̎�ނ𔻕ʂ��܂��B
	Dim filePath
	filePath = GetFilePath(path)
	If IsFile(filePath) Then
		' �t�@�C���������ꍇ�A�J���܂��B
		Editor.FileOpen(filePath)
	Else
		' �t�H���_�������ꍇ�A�ꗗ��\�����܂��B
		Dim list
		list = GetList(path)
		If list <> "" Then
			
			'
			' �w�b�_�𐶐����܂��B
			'
			Dim header
			header = title & vbNewLine
			header = header & separator & vbNewLine
			header = header & path & vbNewLine
			header = header & separator & vbNewLine
			
			'
			' �t�b�^�𐶐����܂��B
			'
			Dim footer
			footer = footer & "* �t�H���_/�t�@�C��/�h���C�u���J��" & vbNewLine
			footer = footer & "    �Ώۍs�ɃJ�[�\�����ړ����ă}�N�������s" & vbNewLine
			footer = footer & vbNewLine
			footer = footer & "* �����[�h" & vbNewLine
			footer = footer & "    3 �s�ڂɃJ�[�\�����ړ����ă}�N�����s" & vbNewLine
			footer = footer & vbNewLine
			footer = footer & "* �߂�" & vbNewLine
			footer = footer & "    �A���h�D(���ɖ߂�)" & vbNewLine
			footer = footer & vbNewLine
			footer = footer & "* �i��" & vbNewLine
			footer = footer & "    ���h�D(��蒼��)" & vbNewLine
			
			If path <> "" Then
				footer = footer & vbNewLine
				footer = footer & "* �t�@�C��/�t�H���_�̍폜" & vbNewLine
				footer = footer & "    �Ώۍs�̐擪�� !del> �Ɠ��͂��A�s���ɃJ�[�\����u�����܂܃}�N�����s" & vbNewLine
				footer = footer & vbNewLine
				footer = footer & "* �t�H���_�쐬" & vbNewLine
				footer = footer & "    �ȉ��̍s�Ƀt�H���_������͂��A�s���ɃJ�[�\����u�����܂܃}�N�����s" & vbNewLine
				footer = footer & "!createfolder>" & vbNewLine
				footer = footer & vbNewLine
				footer = footer & "* �t�@�C���쐬" & vbNewLine
				footer = footer & "    �ȉ��̍s�̃t�@�C��������͂��A�s���ɃJ�[�\����u�����܂܃}�N�����s" & vbNewLine
				footer = footer & "!createfile>" & vbNewLine
			End If
			
			' �o�͂��܂��B
			Dim output
			output = header & list & footer
			Editor.SelectAll()
			Editor.InsText(output)
			Editor.GoFileTop()
			Editor.Down()
			Editor.Down()
			Editor.Down()
			Editor.Down()
			
		End If
	End If
	
End Sub

'
' ���łɃt�@�C���}�l�[�W���ɓ����Ă��邩�ǂ����𒲂ׂ܂��B
'
Function IsFileManager()
	Editor.GoFileTop()
	Editor.GoLineEnd_Sel(0)
	If Editor.GetSelectedString(0) = title Then
		' �G�f�B�^�� 1 �s�ڂ��^�C�g���Ɠ����������ꍇ�́A
		' �t�@�C���}�l�[�W���ɓ����Ă���Ƃ݂Ȃ��B
		IsFileManager = True
	Else
		IsFileManager = False
	End If
	Editor.CancelMode(0)
	Editor.MoveHistPrev()
End Function

'
' �I��͈͂��擾���܂��B
'
Function GetSelection()
	Dim result
	result = Editor.GetSelectedString(0)
	
	' ���s�͍폜
	result = Replace(result, vbCr, "")
	result = Replace(result, vbLf, "")
	
	GetSelection = result
End Function

'
' �J�[�\���ʒu�̍s���擾���܂��B
'
Function GetLine()
	Editor.GoLineTop(0)
	Editor.GoLineEnd_Sel(0)
	GetLine = Editor.GetSelectedString(0)
	Editor.GoLineTop(0)
End Function

'
' ���݂̃p�X���擾���܂��B
'
Function GetCurrentPath()
	' �G�f�B�^�� 3 �s�ڂ�I�����܂��B
	Editor.GoFileTop()
	Editor.Down()
	Editor.Down()
	Editor.GoLineEnd_Sel(0)
	
	' ��������擾���A���`���܂��B
	Dim result
	result = Editor.GetSelectedString(0)
	If result <> "" Then
		If Right(result, 1) <> "\" Then
			result = result & "\"
		End If
	End If
	
	GetCurrentPath = result
	Editor.CancelMode(0)
	Editor.MoveHistPrev()
End Function

'
' �W�J����p�X�����肵�܂��B
'
Function GetPath()
	Dim result
	
	' �t���p�X���ǂ������肵�܂��B
	Dim fullPath
	fullPath = selection
	If fullPath = "" Then
		' �I��͈͂��Ȃ��ꍇ�A���̍s���擾���Ă݂܂��B
		fullPath = GetLine()
	End If
	
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	If fileSystem.FolderExists(fullPath) Then
		' �I��͈͂����̂܂܃t���p�X�Ƃ��ĉ��߂ł��鎞
		result = fullPath
	Else
		' �t���p�X�ł͂Ȃ��Ƃ��A���΃p�X�Ɖ��肵�Ĉȉ��̏������s��
		If isFileManagerMode Then
			' �t�@�C���}�l�[�W�����[�h�̂Ƃ�
			If selection = "" Then
				' �I��͈͂��Ȃ��ꍇ�A���̍s���擾���Ă݂܂��B
				selection = GetLine()
			End If
			result = GetCurrentPath() & selection
		Else
			' �t�@�C���}�l�[�W�����[�h�ł͂Ȃ��Ƃ�
			If selection = "" Then
				result = ""
			Else
				result = selection
			End If
		End If
		
		If result <> "" Then
			If Right(result, 1) <> "\" Then
				result = result & "\"
			End If
		End If
	End If
	
	GetPath = result
End Function

'
' ������ \ ������΂�����������āA�t�@�C���̃t���p�X�̏����ɕϊ����܂��B
'
Function GetFilePath(path)
	Dim filePath
	filePath = path
	' ������ \ �͕s�v
	If Right(filePath, 1) = "\" Then
		filePath = Left(filePath, Len(filePath) - 1)
	End If
	
	GetFilePath = filePath
End Function

'
' �p�X���t�@�C�����ǂ������肵�܂��B
'
Function IsFile(filePath)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	' ���݊m�F
	If filePath <> "" Then
		If fileSystem.FileExists(filePath) Then
			IsFile = True
			Exit Function
		End If
	End If
	
	IsFile = False
End Function

'
' �ꗗ���擾���܂��B
'
Function GetList(path)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	' ���݊m�F
	If path <> "" Then
		If Not fileSystem.FolderExists(path) Then
			Popup("""" & path & """ �͑��݂��܂���B")
			GetList = ""
			Exit Function
		End If
	End If
	
	' �h���C�u�ꗗ
	Dim drives
	drives = GetDriveList(fileSystem)
	drives = drives & separator & vbNewLine
	
	' �t�H���_�ꗗ
	Dim folders
	folders = GetFolderList(fileSystem, path)
	folders = folders & separator & vbNewLine
	
	' �t�@�C���ꗗ
	Dim files
	files = GetFileList(fileSystem, path)
	files = files & separator & vbNewLine
	
	' �ꗗ���������ĕԋp���܂��B
	Dim result
	result = folders & files & drives
	GetList = result
End Function

'
' �h���C�u�ꗗ���擾���܂��B
'
Function GetDriveList(fileSystem)
	Dim drives
	Set drives = fileSystem.Drives
	
	Dim result
	
	' �e�h���C�u���������܂��B
	Dim drive
	For Each drive In drives
		result = result & drive.DriveLetter & ":\" & vbNewLine
	Next
	
	If result = "" Then
		result = "(�h���C�u����)" & vbNewLine
	End If
	
	GetDriveList = result
End Function

'
' �t�H���_�ꗗ���擾���܂��B
'
Function GetFolderList(fileSystem, path)
	If path = "" Then
		GetFolderList = ""
	Else
		Dim folder
		Set folder = fileSystem.GetFolder(path)
		
		Dim result
		
		' �e�t�H���_���������܂��B
		Dim foldername
		For Each folderName In folder.Subfolders
			folderName = Mid(folderName, Len(path) + 1)
			result = result & folderName & vbNewLine
		Next
	End If
	
	If result = "" Then
		result = "(�t�H���_����)" & vbNewLine
	End If
	
	GetFolderList = result
End Function

'
' �t�@�C���ꗗ���擾���܂��B
'
Function GetFileList(fileSystem, path)
	If path = "" Then
		GetFileList = ""
	Else
		Dim folder
		Set folder = fileSystem.GetFolder(path)
		
		Dim result
		
		' �e�t�@�C�����������܂��B
		For Each fileName In folder.Files
			fileName = Mid(fileName, Len(path) + 1)
			result = result & fileName & vbNewLine
		Next
	End If
	
	If result = "" Then
		result = "(�t�@�C������)" & vbNewLine
	End If
	
	GetFileList = result
End Function

'
' �R�}���h�����s���܂��B
'
Sub DoCommand(commandLine)
	Dim commands
	commands = Split(commandLine, ">")
	Select Case commands(0)
		Case "!createfolder" : MakeDirectory(commands(1))
		Case "!createfile" : MakeFile(commands(1))
		Case "!del" : Delete(commands(1))
	End Select
End Sub

'
' �f�B���N�g�����쐬���܂��B
'
Sub MakeDirectory(directoryName)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	Dim path
	path = GetCurrentPath() & directoryName
	
	If Not fileSystem.FolderExists(path) And _
	   Not fileSystem.FileExists(path) Then
		fileSystem.CreateFolder(path)
		Popup("""" & path & """ ���쐬���܂����B�����[�h���ĉ������B")
		Editor.GoFileTop()
		Editor.Down()
		Editor.Down()
	Else
		Popup("""" & path & """ �͂��łɑ��݂��Ă��܂��B")
	End If
End Sub

'
' �t�@�C�����쐬���܂��B
'
Sub MakeFile(fileName)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	Dim path
	path = GetCurrentPath() & fileName
	
	If Not fileSystem.FolderExists(path) And _
	   Not fileSystem.FileExists(path) Then
		Call fileSystem.CreateTextFile(path, False)
		Popup("""" & path & """ ���쐬���܂����B�����[�h���ĉ������B")
		Editor.GoFileTop()
		Editor.Down()
		Editor.Down()
	Else
		Popup("""" & path & """ �͂��łɑ��݂��Ă��܂��B")
	End If
End Sub

'
' �t�H���_�܂��̓t�@�C�����폜���܂��B
'
Sub Delete(name)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	Dim path
	path = GetCurrentPath() & name
	
	Dim isExistsFolder
	isExistsFolder = fileSystem.FolderExists(path)
	
	Dim isExistsFile
	isExistsFile = fileSystem.FileExists(path)
	
	If isExistsFolder Then
		Call fileSystem.DeleteFolder(path, False)
		Popup("""" & path & """ ���폜���܂����B�����[�h���ĉ������B")
		Editor.GoFileTop()
		Editor.Down()
		Editor.Down()
	ElseIf isExistsFile Then
		Call fileSystem.DeleteFile(path, False)
		Popup("""" & path & """ ���폜���܂����B�����[�h���ĉ������B")
		Editor.GoFileTop()
		Editor.Down()
		Editor.Down()
	Else
		Popup("""" & path & """ ��������܂���B")
	End If
End Sub

'
' ���b�Z�[�W���|�b�v�A�b�v�\�����܂��B
'
Function Popup(message)
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    Call shell.Popup(message, 0, "", 0)
End Function

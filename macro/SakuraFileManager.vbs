'
' Sakura File Manager 1.0 beta 1
'
' (c)2005 ぷーる
' webmaster@muumoo.jp
' http://muumoo.jp/
'

'
' 設定
'
' タイトル
Dim title
title = "Sakura File Manager"
' バージョン
Dim version
version = "1.0 beta 1"
' 区切り線
Dim separator
separator = "----------------"
' 選択領域
Dim selection
selection = GetSelection()
' すでにファイルマネージャに入っているか否か
Dim isFileManagerMode
isFileManagerMode = IsFileManager()

' 実行
Main()

'
' メイン
'
Sub Main
	
	' コマンドかどうか判定
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
	
	' 対象パスを取得します。
	Dim path
	path = GetPath()
	
	' パスの種類を判別します。
	Dim filePath
	filePath = GetFilePath(path)
	If IsFile(filePath) Then
		' ファイルだった場合、開きます。
		Editor.FileOpen(filePath)
	Else
		' フォルダだった場合、一覧を表示します。
		Dim list
		list = GetList(path)
		If list <> "" Then
			
			'
			' ヘッダを生成します。
			'
			Dim header
			header = title & vbNewLine
			header = header & separator & vbNewLine
			header = header & path & vbNewLine
			header = header & separator & vbNewLine
			
			'
			' フッタを生成します。
			'
			Dim footer
			footer = footer & "* フォルダ/ファイル/ドライブを開く" & vbNewLine
			footer = footer & "    対象行にカーソルを移動してマクロを実行" & vbNewLine
			footer = footer & vbNewLine
			footer = footer & "* リロード" & vbNewLine
			footer = footer & "    3 行目にカーソルを移動してマクロ実行" & vbNewLine
			footer = footer & vbNewLine
			footer = footer & "* 戻る" & vbNewLine
			footer = footer & "    アンドゥ(元に戻す)" & vbNewLine
			footer = footer & vbNewLine
			footer = footer & "* 進む" & vbNewLine
			footer = footer & "    リドゥ(やり直し)" & vbNewLine
			
			If path <> "" Then
				footer = footer & vbNewLine
				footer = footer & "* ファイル/フォルダの削除" & vbNewLine
				footer = footer & "    対象行の先頭に !del> と入力し、行内にカーソルを置いたままマクロ実行" & vbNewLine
				footer = footer & vbNewLine
				footer = footer & "* フォルダ作成" & vbNewLine
				footer = footer & "    以下の行にフォルダ名を入力し、行内にカーソルを置いたままマクロ実行" & vbNewLine
				footer = footer & "!createfolder>" & vbNewLine
				footer = footer & vbNewLine
				footer = footer & "* ファイル作成" & vbNewLine
				footer = footer & "    以下の行のファイル名を入力し、行内にカーソルを置いたままマクロ実行" & vbNewLine
				footer = footer & "!createfile>" & vbNewLine
			End If
			
			' 出力します。
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
' すでにファイルマネージャに入っているかどうかを調べます。
'
Function IsFileManager()
	Editor.GoFileTop()
	Editor.GoLineEnd_Sel(0)
	If Editor.GetSelectedString(0) = title Then
		' エディタの 1 行目がタイトルと等しかった場合は、
		' ファイルマネージャに入っているとみなす。
		IsFileManager = True
	Else
		IsFileManager = False
	End If
	Editor.CancelMode(0)
	Editor.MoveHistPrev()
End Function

'
' 選択範囲を取得します。
'
Function GetSelection()
	Dim result
	result = Editor.GetSelectedString(0)
	
	' 改行は削除
	result = Replace(result, vbCr, "")
	result = Replace(result, vbLf, "")
	
	GetSelection = result
End Function

'
' カーソル位置の行を取得します。
'
Function GetLine()
	Editor.GoLineTop(0)
	Editor.GoLineEnd_Sel(0)
	GetLine = Editor.GetSelectedString(0)
	Editor.GoLineTop(0)
End Function

'
' 現在のパスを取得します。
'
Function GetCurrentPath()
	' エディタの 3 行目を選択します。
	Editor.GoFileTop()
	Editor.Down()
	Editor.Down()
	Editor.GoLineEnd_Sel(0)
	
	' 文字列を取得し、整形します。
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
' 展開するパスを決定します。
'
Function GetPath()
	Dim result
	
	' フルパスかどうか判定します。
	Dim fullPath
	fullPath = selection
	If fullPath = "" Then
		' 選択範囲がない場合、その行を取得してみます。
		fullPath = GetLine()
	End If
	
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	If fileSystem.FolderExists(fullPath) Then
		' 選択範囲がそのままフルパスとして解釈できる時
		result = fullPath
	Else
		' フルパスではないとき、相対パスと仮定して以下の処理を行う
		If isFileManagerMode Then
			' ファイルマネージャモードのとき
			If selection = "" Then
				' 選択範囲がない場合、その行を取得してみます。
				selection = GetLine()
			End If
			result = GetCurrentPath() & selection
		Else
			' ファイルマネージャモードではないとき
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
' 末尾に \ があればそれを除去して、ファイルのフルパスの書式に変換します。
'
Function GetFilePath(path)
	Dim filePath
	filePath = path
	' 末尾の \ は不要
	If Right(filePath, 1) = "\" Then
		filePath = Left(filePath, Len(filePath) - 1)
	End If
	
	GetFilePath = filePath
End Function

'
' パスがファイルかどうか判定します。
'
Function IsFile(filePath)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	' 存在確認
	If filePath <> "" Then
		If fileSystem.FileExists(filePath) Then
			IsFile = True
			Exit Function
		End If
	End If
	
	IsFile = False
End Function

'
' 一覧を取得します。
'
Function GetList(path)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	' 存在確認
	If path <> "" Then
		If Not fileSystem.FolderExists(path) Then
			Popup("""" & path & """ は存在しません。")
			GetList = ""
			Exit Function
		End If
	End If
	
	' ドライブ一覧
	Dim drives
	drives = GetDriveList(fileSystem)
	drives = drives & separator & vbNewLine
	
	' フォルダ一覧
	Dim folders
	folders = GetFolderList(fileSystem, path)
	folders = folders & separator & vbNewLine
	
	' ファイル一覧
	Dim files
	files = GetFileList(fileSystem, path)
	files = files & separator & vbNewLine
	
	' 一覧を結合して返却します。
	Dim result
	result = folders & files & drives
	GetList = result
End Function

'
' ドライブ一覧を取得します。
'
Function GetDriveList(fileSystem)
	Dim drives
	Set drives = fileSystem.Drives
	
	Dim result
	
	' 各ドライブを処理します。
	Dim drive
	For Each drive In drives
		result = result & drive.DriveLetter & ":\" & vbNewLine
	Next
	
	If result = "" Then
		result = "(ドライブ無し)" & vbNewLine
	End If
	
	GetDriveList = result
End Function

'
' フォルダ一覧を取得します。
'
Function GetFolderList(fileSystem, path)
	If path = "" Then
		GetFolderList = ""
	Else
		Dim folder
		Set folder = fileSystem.GetFolder(path)
		
		Dim result
		
		' 各フォルダを処理します。
		Dim foldername
		For Each folderName In folder.Subfolders
			folderName = Mid(folderName, Len(path) + 1)
			result = result & folderName & vbNewLine
		Next
	End If
	
	If result = "" Then
		result = "(フォルダ無し)" & vbNewLine
	End If
	
	GetFolderList = result
End Function

'
' ファイル一覧を取得します。
'
Function GetFileList(fileSystem, path)
	If path = "" Then
		GetFileList = ""
	Else
		Dim folder
		Set folder = fileSystem.GetFolder(path)
		
		Dim result
		
		' 各ファイルを処理します。
		For Each fileName In folder.Files
			fileName = Mid(fileName, Len(path) + 1)
			result = result & fileName & vbNewLine
		Next
	End If
	
	If result = "" Then
		result = "(ファイル無し)" & vbNewLine
	End If
	
	GetFileList = result
End Function

'
' コマンドを実行します。
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
' ディレクトリを作成します。
'
Sub MakeDirectory(directoryName)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	Dim path
	path = GetCurrentPath() & directoryName
	
	If Not fileSystem.FolderExists(path) And _
	   Not fileSystem.FileExists(path) Then
		fileSystem.CreateFolder(path)
		Popup("""" & path & """ を作成しました。リロードして下さい。")
		Editor.GoFileTop()
		Editor.Down()
		Editor.Down()
	Else
		Popup("""" & path & """ はすでに存在しています。")
	End If
End Sub

'
' ファイルを作成します。
'
Sub MakeFile(fileName)
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	Dim path
	path = GetCurrentPath() & fileName
	
	If Not fileSystem.FolderExists(path) And _
	   Not fileSystem.FileExists(path) Then
		Call fileSystem.CreateTextFile(path, False)
		Popup("""" & path & """ を作成しました。リロードして下さい。")
		Editor.GoFileTop()
		Editor.Down()
		Editor.Down()
	Else
		Popup("""" & path & """ はすでに存在しています。")
	End If
End Sub

'
' フォルダまたはファイルを削除します。
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
		Popup("""" & path & """ を削除しました。リロードして下さい。")
		Editor.GoFileTop()
		Editor.Down()
		Editor.Down()
	ElseIf isExistsFile Then
		Call fileSystem.DeleteFile(path, False)
		Popup("""" & path & """ を削除しました。リロードして下さい。")
		Editor.GoFileTop()
		Editor.Down()
		Editor.Down()
	Else
		Popup("""" & path & """ が見つかりません。")
	End If
End Sub

'
' メッセージをポップアップ表示します。
'
Function Popup(message)
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    Call shell.Popup(message, 0, "", 0)
End Function

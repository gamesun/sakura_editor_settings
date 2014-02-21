' 選択文字列/カーソル位置単語をGoogleで検索
Dim keyword

keyword = GetSelectedString(0)
If keyword = "" Then
    keyword = GetSelectedString(0)
End If

Dim sh
Set sh = CreateObject("WScript.Shell")
Call sh.Run("http://www.google.co.jp/search?ie=Shift_JIS&hl=ja&lr=&num=50&q=" & keyword)

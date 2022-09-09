Option Explicit

Dim fs, fd, fPath, f, fTxt, fStr

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd = fs.GetFolder(fPath & "\TestData")

For Each f In fd.Files

	'テキストファイルを読み取り専用で開く
	Set fTxt = f.OpenAsTextStream(1, 0)

	'テキストファイル全体を読み込む
	fStr = fTxt.ReadAll

	'「埼玉県」が含まれていたら表示する
	If InStr(fStr, "埼玉県") > 0 Then MsgBox fStr,, f.Name
	
	'テキストファイルを閉じる
	fTxt.Close

	Set fTxt = Nothing
Next

Set fd = Nothing
Set fs = Nothing
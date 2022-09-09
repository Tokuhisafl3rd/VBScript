Option Explicit

Dim fs, fPath, fd, f, fns

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd = fs.GetFolder(fPath & "\TestData")

'対象フォルダー内の全ファイルについて繰り返す
For Each f In fd.Files

	'ファイル名を追加する
	fns = fns & f.Name & vbCr
Next

'最終行に全ファイル数を追加する
fns = fns & "以上" & fd.Files.Count & "ファイル"

'全ファイル名を表示する
MsgBox fns,, "全ファイル名表示"

Set fd = Nothing
Set fs = Nothing
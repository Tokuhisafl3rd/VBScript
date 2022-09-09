Option Explicit

Dim fs, fd1, fPath, fd2, f, cnt

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd1 = fs.GetFolder(fPath & "\TestData")

'「Items」フォルダーを作成する
Set fd2 = fs.CreateFolder(fd1.Path & "\Items")

cnt = 0

For Each f In fd1.Files

	'「item」が含まれていたら
	If InStr(f.Name, "item") > 0 Then

		'ファイルを移動する
		f.Move fd2.Path & "\"

		'移動した数を記録する
		cnt = cnt + 1
	End If
Next

MsgBox cnt & "個のファイルを移動しました",, "移動完了"

Set fd1 = Nothing
Set fd2 = Nothing
Set fs = Nothing
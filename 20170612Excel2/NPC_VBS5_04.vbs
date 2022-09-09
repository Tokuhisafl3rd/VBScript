Option Explicit

Dim fs, fPath, fd, f, cnt

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd = fs.GetFolder(fPath & "\TestData")

cnt = 0

For Each f In fd.Files

	'「user」が含まれていたら
	If InStr(f.Name, "user") > 0 Then

		'ファイル名を変更する
		f.Name = Replace(f.Name, "user", "member")

		'変更した数を記録する
		cnt = cnt + 1
	End If
Next

MsgBox cnt & "個のファイル名を変更しました",, "変更完了"

Set fd = Nothing
Set fs = Nothing
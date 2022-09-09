Option Explicit

Dim fs, fd, fPath, f, fTxt, fStr, cnt

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd = fs.GetFolder(fPath & "\TestData")

cnt = 0

For Each f In fd.Files

	Set fTxt = f.OpenAsTextStream(1, 0)

	fStr = fTxt.ReadAll

	fTxt.Close

	'テキスト中に「番号：」があれば
	If InStr(fStr, "番号：") > 0 Then

		'同じテキストファイルを上書き用に開く
		Set fTxt = f.OpenAsTextStream(2, 0)

		'データを変更して書き込む
		fTxt.Write Replace(fStr, "番号：", "No.")
		
		fTxt.Close

		cnt = cnt + 1
	End If

	Set fTxt = Nothing
Next

MsgBox cnt & "個のファイルを変更しました",, "変更完了"

Set fd = Nothing
Set fs = Nothing
Option Explicit

Dim fs, fPath, fd, f, cnt

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd = fs.GetFolder(fPath & "\TestData")

cnt = 0

For Each f In fd.Files

	'�uuser�v���܂܂�Ă�����
	If InStr(f.Name, "user") > 0 Then

		'�t�@�C������ύX����
		f.Name = Replace(f.Name, "user", "member")

		'�ύX���������L�^����
		cnt = cnt + 1
	End If
Next

MsgBox cnt & "�̃t�@�C������ύX���܂���",, "�ύX����"

Set fd = Nothing
Set fs = Nothing
Option Explicit

Dim fs, fd1, fPath, fd2, f, cnt

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd1 = fs.GetFolder(fPath & "\TestData")

'�uItems�v�t�H���_�[���쐬����
Set fd2 = fs.CreateFolder(fd1.Path & "\Items")

cnt = 0

For Each f In fd1.Files

	'�uitem�v���܂܂�Ă�����
	If InStr(f.Name, "item") > 0 Then

		'�t�@�C�����ړ�����
		f.Move fd2.Path & "\"

		'�ړ����������L�^����
		cnt = cnt + 1
	End If
Next

MsgBox cnt & "�̃t�@�C�����ړ����܂���",, "�ړ�����"

Set fd1 = Nothing
Set fd2 = Nothing
Set fs = Nothing
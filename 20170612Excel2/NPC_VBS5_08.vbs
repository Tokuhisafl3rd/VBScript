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

	'�e�L�X�g���Ɂu�ԍ��F�v�������
	If InStr(fStr, "�ԍ��F") > 0 Then

		'�����e�L�X�g�t�@�C�����㏑���p�ɊJ��
		Set fTxt = f.OpenAsTextStream(2, 0)

		'�f�[�^��ύX���ď�������
		fTxt.Write Replace(fStr, "�ԍ��F", "No.")
		
		fTxt.Close

		cnt = cnt + 1
	End If

	Set fTxt = Nothing
Next

MsgBox cnt & "�̃t�@�C����ύX���܂���",, "�ύX����"

Set fd = Nothing
Set fs = Nothing
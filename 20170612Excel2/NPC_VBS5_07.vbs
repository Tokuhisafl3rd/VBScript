Option Explicit

Dim fs, fd, fPath, f, fTxt, fStr

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd = fs.GetFolder(fPath & "\TestData")

For Each f In fd.Files

	'�e�L�X�g�t�@�C����ǂݎ���p�ŊJ��
	Set fTxt = f.OpenAsTextStream(1, 0)

	'�e�L�X�g�t�@�C���S�̂�ǂݍ���
	fStr = fTxt.ReadAll

	'�u��ʌ��v���܂܂�Ă�����\������
	If InStr(fStr, "��ʌ�") > 0 Then MsgBox fStr,, f.Name
	
	'�e�L�X�g�t�@�C�������
	fTxt.Close

	Set fTxt = Nothing
Next

Set fd = Nothing
Set fs = Nothing
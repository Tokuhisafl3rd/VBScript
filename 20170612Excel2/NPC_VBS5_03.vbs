Option Explicit

Dim fs, fPath, fd, f, fns

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

Set fd = fs.GetFolder(fPath & "\TestData")

'�Ώۃt�H���_�[���̑S�t�@�C���ɂ��ČJ��Ԃ�
For Each f In fd.Files

	'�t�@�C������ǉ�����
	fns = fns & f.Name & vbCr
Next

'�ŏI�s�ɑS�t�@�C������ǉ�����
fns = fns & "�ȏ�" & fd.Files.Count & "�t�@�C��"

'�S�t�@�C������\������
MsgBox fns,, "�S�t�@�C�����\��"

Set fd = Nothing
Set fs = Nothing
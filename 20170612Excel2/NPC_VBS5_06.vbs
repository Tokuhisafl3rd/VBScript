Option Explicit

Dim fs, fPath, fTxt

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

'�V�����e�L�X�g�t�@�C�����쐬����
Set fTxt = fs.CreateTextFile(fPath & "\TestData\member06.txt")

'�f�[�^����������
fTxt.WriteLine "�ԍ��F06"
fTxt.WriteLine "�쓇���ގq"
fTxt.WriteLine "41��"
fTxt.WriteLine "�����s�䓌��"
		
fTxt.Close

Set fTxt = Nothing
Set fs = Nothing
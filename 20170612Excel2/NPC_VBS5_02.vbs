Option Explicit

Dim fs, fPath, fd, fNum

Set fs = CreateObject("Scripting.FileSystemObject")

'���̃t�@�C���̐e�t�H���_�[�̃p�X���擾����
fPath = fs.GetParentFolderName(WScript.ScriptFullName)

'�uTestData�v�t�H���_�[��Folder�I�u�W�F�N�g�Ƃ��Ď擾����
Set fd = fs.GetFolder(fPath & "\TestData")

'�t�H���_�[���̃t�@�C�����𒲂ׂ�
fNum = fd.Files.Count

MsgBox fNum & "�t�@�C���ł�",, "TestData"

Set fd = Nothing
Set fs = Nothing
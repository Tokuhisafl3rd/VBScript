Option Explicit

Dim fs

'FileSystemObject�I�u�W�F�N�g���쐬
Set fs = CreateObject("Scripting.FileSystemObject")

'����̃t�H���_�[�̗L���𒲂ׂ�
If fs.FolderExists("C:\Users\����\Documents\TestData") Then

	MsgBox "�w��̃t�H���_�[�͑��݂��܂�",, "���݊m�F"
Else
	MsgBox "�w��̃t�H���_�[�͑��݂��܂���",, "���݊m�F"
End If

'�g�p�����I�u�W�F�N�g�ϐ����������
Set fs = Nothing
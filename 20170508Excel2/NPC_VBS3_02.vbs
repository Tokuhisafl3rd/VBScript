Option Explicit

Dim pw

Do
	'�p�X���[�h�����
	pw = InputBox("�p�X���[�h����͂��Ă�������","�p�X���[�h") 

	'�u�L�����Z���v���ǂ����𔻒�
	If IsEmpty(pw) Then
		WScript.Quit

	'�𓚂̔���
	ElseIf pw = "VBScript" Then
		Exit Do
	End If
Loop

MsgBox "�p�X���[�h��F�؂��܂���",,"�F��"
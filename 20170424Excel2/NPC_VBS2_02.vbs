Option Explicit

Dim ryo, pnum

ryo = InputBox("�w��������͂��Ă�������")

If ryo >= 5 Then

	'���i�ԍ��̓���
	pnum = InputBox("���i�ԍ�����͂��Ă�������")

	'�w�����e�̕\��
	MsgBox "���i�ԍ��F" & pnum & vbCr & _
		"�w�����z�F" & 250 * ryo & "�~" 
End If
Option Explicit

Dim pnum, ryo

'���i�ԍ��̓���
pnum = InputBox("���i�ԍ�����͂��Ă�������")

'���i�ԍ��̔���
If pnum >= 101 And pnum <= 105 Then

	ryo = InputBox("�w��������͂��Ă�������")

	MsgBox "���i�ԍ��F" & pnum & vbCr & _
		"�w�����z�F" & 250 * ryo & "�~" 
Else

	'���i�ԍ����s�K�؂ȏꍇ
	MsgBox "���̏��i�͎�舵���Ă��܂���"

End If
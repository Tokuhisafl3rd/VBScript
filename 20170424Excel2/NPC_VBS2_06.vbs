Option Explicit

Dim pnum, ryo

pnum = InputBox("���i�ԍ�����͂��Ă�������")

If pnum >= 101 And pnum <= 105 Then

	'�w�����̓���
	ryo = InputBox("�w��������͂��Ă�������")

	'�w�����̔���
	If ryo >= 5 Then

		MsgBox "���i�ԍ��F" & pnum & vbCr & _
			"�w�����z�F" & 250 * ryo & "�~" 
	Else
		MsgBox "5�ȏ�ł��w����������"
	End If
Else
	MsgBox "���̏��i�͎�舵���Ă��܂���"

End If
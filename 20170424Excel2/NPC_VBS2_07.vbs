Option Explicit

Dim pnum, ryo

pnum = InputBox("���i�ԍ�����͂��Ă�������")

'�u�L�����Z���v�������ꂽ�ꍇ
If pnum = "" Then

	'�����I��
	WScript.Quit

ElseIf pnum >= 101 And pnum <= 105 Then

	ryo = InputBox("�w��������͂��Ă�������")

	'�u�L�����Z���v�������ꂽ�ꍇ
	If ryo = "" Then

		'�����I��
		WScript.Quit 
	
	ElseIf ryo >= 5 Then

		MsgBox "���i�ԍ��F" & pnum & vbCr & _
			"�w�����z�F" & 250 * ryo & "�~" 
	Else
		MsgBox "5�ȏ�ł��w����������"
	End If
Else
	MsgBox "���̏��i�͎�舵���Ă��܂���"

End If
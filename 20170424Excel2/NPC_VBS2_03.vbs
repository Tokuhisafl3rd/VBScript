Option Explicit

Dim ryo

ryo = InputBox("�w��������͂��Ă�������")

If ryo >= 10 Then

	'�������i�Ōv�Z
	MsgBox "�w�����z�F" & 240 * ryo & "�~" 
Else

	'�ʏ퉿�i�Ōv�Z
	MsgBox "�w�����z�F" & 250 * ryo & "�~"
End If
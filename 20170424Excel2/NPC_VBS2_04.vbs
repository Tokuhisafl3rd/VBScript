Option Explicit

Dim ryo

ryo = InputBox("�w��������͂��Ă�������")

If ryo < 5 Then

	'���Ȃ��ꍇ�̃��b�Z�[�W
	MsgBox "5�ȏ�ł��w����������"
	
ElseIf ryo >= 10 Then

	MsgBox "�w�����z�F" & 240 * ryo & "�~" 
Else

	MsgBox "�w�����z�F" & 250 * ryo & "�~"
End If
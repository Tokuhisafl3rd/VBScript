Option Explicit

Dim pnum, tanka, ryo

pnum = InputBox("���i�ԍ�����͂��Ă�������")

If pnum = "" Then WScript.Quit

'pnum�̒l��]��
Select Case pnum

'���i�ԍ��ɉ������P�����擾
Case "101" tanka = 200
Case "102" tanka = 220
Case "103" tanka = 240
Case "104" tanka = 250
Case "105" tanka = 280

'���o�^�̏��i�ԍ�
Case Else

	MsgBox "���i�ԍ����s�K�؂ł�"
	WScript.Quit

End Select

ryo = InputBox("�w��������͂��Ă�������")

If ryo = "" Or Not IsNumeric(ryo) Then

	WScript.Quit

ElseIf ryo >= 5 Then

	MsgBox "���i�ԍ��F" & pnum & vbCr & _
		"�w�����z�F" & tanka * ryo & "�~" 

Else
	MsgBox "5�ȏ�ł��w����������"
End If
Option Explicit

Dim num1, num2

num1 = InputBox("���肵�������l����͂��Ă�������", "�f������")

If IsEmpty(num1) Then
	WScript.Quit

ElseIf IsNumeric(num1) Then

	num1 = Int(num1)

	If num1 < 2 Then
		MsgBox num1 & "�͑f���ł͂���܂���",, "�f������"
		WScript.Quit
	End If
Else
	MsgBox "���l����͂��Ă�������",, "�f������"
	WScript.Quit
End If

'2���琔�l�̐��̕������̒l�܂ŌJ��Ԃ�
For num2 = 2 To Int(Sqr(num1))

	'����؂�邩�ǂ����𔻒�
	If num1 Mod num2 = 0 Then
		MsgBox num1 & "�͑f���ł͂���܂���",, "�f������"
		WScript.Quit
	End If
Next

MsgBox num1 & "�͑f���ł�",, "�f������"

Option Explicit

Dim num1, num2, num3

num1 = InputBox("0�`15�̐�������͂��Ă�������", "�l�̓���")

'�u�L�����Z���v���ǂ������`�F�b�N
If IsEmpty(num1) Then
	WScript.Quit

'���l���ǂ������`�F�b�N
ElseIf IsNumeric(num1) Then

	'�����_�ȉ��̒[����؂�̂�
	num1 = Int(num1)

	'�w��͈͓̔����ǂ������`�F�b�N
	If num1 > 15 Or num1 < 0 Then
		MsgBox "�s�K�؂Ȑ��l�ł�"
		WScript.Quit
	End If
Else
	MsgBox "���l����͂��Ă�������"
	WScript.Quit
End If

num2 = 1

For num3 = 2 To num1
	num2 = num2 * num3
Next

MsgBox num1 & "�̊K��F" & num2, , "�K��̒l"

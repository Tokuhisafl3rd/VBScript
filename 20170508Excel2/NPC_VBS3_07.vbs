Option Explicit

Dim num1, num2, num3

num1 = InputBox("0�`15�̐�������͂��Ă�������", "�l�̓���")

num2 = 1

For num3 = 2 To num1
	num2 = num2 * num3
Next

MsgBox num1 & "�̊K��F" & num2, , "�K��̒l"

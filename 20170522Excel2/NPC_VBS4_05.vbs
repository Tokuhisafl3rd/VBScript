Option Explicit

Dim num1, num2

num1 = InputBox("1�Ԗڂ̐�������͂��Ă�������", "���l����")
Call chkInput(num1)

num2 = InputBox("2�Ԗڂ̐�������͂��Ă�������", "���l����")
Call chkInput(num2)

MsgBox num1 & "��" & num2 & "�̍ő���񐔂�" & getGCD(num1, num2) & "�ł�",, _
	"���ʕ\��"

Sub chkInput(inum)
	If IsEmpty(inum) Then
		MsgBox "�����𒆎~���܂�",, "���~"
		WScript.Quit
	ElseIf Not IsNumeric(inum) Then
		MsgBox "���͒l���s�K�؂ł�",, "���~"
		WScript.Quit
	ElseIf inum < 1 Or inum - Int(inum) > 0 Then
		MsgBox "���̐�������͂��Ă�������",, "���~"
		WScript.Quit
	End If
End Sub

'�ő���񐔂����߂�֐����ċA�@�ō쐬
Function getGCD(numA, numB)

	If numA Mod numB = 0 Then
		getGCD = numB
	Else
		getGCD = getGCD(numB, numA Mod numB)
	End If

End Function
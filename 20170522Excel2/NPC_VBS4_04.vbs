Option Explicit

Dim num1, num2

num1 = InputBox("1�Ԗڂ̐�������͂��Ă�������", "���l����")
Call chkInput(num1)

num2 = InputBox("2�Ԗڂ̐�������͂��Ă�������", "���l����")
Call chkInput(num2)

'2�̐����̍ő���񐔂�\��
MsgBox num1 & "��" & num2 & "�̍ő���񐔂�" & getGCD(num1, num2) & "�ł�",, _
	"���ʕ\��"

'�m�F�p�T�u���[�`��
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

'�ő���񐔂����߂�֐�
Function getGCD(numA, numB)
	Dim a, b, r
	
	a = numA
	b = numB

	Do
		r = a Mod b
		If r = 0 Then
			getGCD = b
			Exit Function
		Else
			a = b
			b = r
		End If
	Loop

End Function
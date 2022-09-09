Option Explicit

Dim num1, num2

num1 = InputBox("判定したい数値を入力してください", "素数判定")

If IsEmpty(num1) Then
	WScript.Quit

ElseIf IsNumeric(num1) Then

	num1 = Int(num1)

	If num1 < 2 Then
		MsgBox num1 & "は素数ではありません",, "素数判定"
		WScript.Quit
	End If
Else
	MsgBox "数値を入力してください",, "素数判定"
	WScript.Quit
End If

'2から数値の正の平方根の値まで繰り返し
For num2 = 2 To Int(Sqr(num1))

	'割り切れるかどうかを判定
	If num1 Mod num2 = 0 Then
		MsgBox num1 & "は素数ではありません",, "素数判定"
		WScript.Quit
	End If
Next

MsgBox num1 & "は素数です",, "素数判定"

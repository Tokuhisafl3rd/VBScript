Option Explicit

Dim num1, num2, num3

num1 = InputBox("0～15の整数を入力してください", "値の入力")

'「キャンセル」かどうかをチェック
If IsEmpty(num1) Then
	WScript.Quit

'数値かどうかをチェック
ElseIf IsNumeric(num1) Then

	'小数点以下の端数を切り捨て
	num1 = Int(num1)

	'指定の範囲内かどうかをチェック
	If num1 > 15 Or num1 < 0 Then
		MsgBox "不適切な数値です"
		WScript.Quit
	End If
Else
	MsgBox "数値を入力してください"
	WScript.Quit
End If

num2 = 1

For num3 = 2 To num1
	num2 = num2 * num3
Next

MsgBox num1 & "の階乗：" & num2, , "階乗の値"

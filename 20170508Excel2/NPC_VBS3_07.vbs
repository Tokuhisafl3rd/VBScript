Option Explicit

Dim num1, num2, num3

num1 = InputBox("0〜15の整数を入力してください", "値の入力")

num2 = 1

For num3 = 2 To num1
	num2 = num2 * num3
Next

MsgBox num1 & "の階乗：" & num2, , "階乗の値"

Option Explicit

Dim num1, num2

num1 = InputBox("1番目の整数を入力してください", "数値入力")
Call chkInput(num1)

num2 = InputBox("2番目の整数を入力してください", "数値入力")
Call chkInput(num2)

MsgBox num1 & "と" & num2 & "の最大公約数は" & getGCD(num1, num2) & "です",, _
	"結果表示"

Sub chkInput(inum)
	If IsEmpty(inum) Then
		MsgBox "処理を中止します",, "中止"
		WScript.Quit
	ElseIf Not IsNumeric(inum) Then
		MsgBox "入力値が不適切です",, "中止"
		WScript.Quit
	ElseIf inum < 1 Or inum - Int(inum) > 0 Then
		MsgBox "正の整数を入力してください",, "中止"
		WScript.Quit
	End If
End Sub

'最大公約数を求める関数を再帰法で作成
Function getGCD(numA, numB)

	If numA Mod numB = 0 Then
		getGCD = numB
	Else
		getGCD = getGCD(numB, numA Mod numB)
	End If

End Function
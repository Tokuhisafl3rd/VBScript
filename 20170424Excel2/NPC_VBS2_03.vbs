Option Explicit

Dim ryo

ryo = InputBox("購入数を入力してください")

If ryo >= 10 Then

	'割引価格で計算
	MsgBox "購入金額：" & 240 * ryo & "円" 
Else

	'通常価格で計算
	MsgBox "購入金額：" & 250 * ryo & "円"
End If
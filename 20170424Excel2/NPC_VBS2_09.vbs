Option Explicit

Dim pnum, tanka, ryo

pnum = InputBox("商品番号を入力してください")

If pnum = "" Then WScript.Quit

'pnumの値を評価
Select Case pnum

'商品番号に応じた単価を取得
Case "101" tanka = 200
Case "102" tanka = 220
Case "103" tanka = 240
Case "104" tanka = 250
Case "105" tanka = 280

'未登録の商品番号
Case Else

	MsgBox "商品番号が不適切です"
	WScript.Quit

End Select

ryo = InputBox("購入数を入力してください")

If ryo = "" Or Not IsNumeric(ryo) Then

	WScript.Quit

ElseIf ryo >= 5 Then

	MsgBox "商品番号：" & pnum & vbCr & _
		"購入金額：" & tanka * ryo & "円" 

Else
	MsgBox "5個以上でご購入ください"
End If
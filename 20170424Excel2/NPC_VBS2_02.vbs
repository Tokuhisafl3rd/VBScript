Option Explicit

Dim ryo, pnum

ryo = InputBox("購入数を入力してください")

If ryo >= 5 Then

	'商品番号の入力
	pnum = InputBox("商品番号を入力してください")

	'購入内容の表示
	MsgBox "商品番号：" & pnum & vbCr & _
		"購入金額：" & 250 * ryo & "円" 
End If
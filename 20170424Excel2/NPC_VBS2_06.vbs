Option Explicit

Dim pnum, ryo

pnum = InputBox("商品番号を入力してください")

If pnum >= 101 And pnum <= 105 Then

	'購入数の入力
	ryo = InputBox("購入数を入力してください")

	'購入数の判定
	If ryo >= 5 Then

		MsgBox "商品番号：" & pnum & vbCr & _
			"購入金額：" & 250 * ryo & "円" 
	Else
		MsgBox "5個以上でご購入ください"
	End If
Else
	MsgBox "その商品は取り扱っていません"

End If
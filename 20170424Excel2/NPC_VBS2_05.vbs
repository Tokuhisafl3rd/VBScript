Option Explicit

Dim pnum, ryo

'商品番号の入力
pnum = InputBox("商品番号を入力してください")

'商品番号の判定
If pnum >= 101 And pnum <= 105 Then

	ryo = InputBox("購入数を入力してください")

	MsgBox "商品番号：" & pnum & vbCr & _
		"購入金額：" & 250 * ryo & "円" 
Else

	'商品番号が不適切な場合
	MsgBox "その商品は取り扱っていません"

End If
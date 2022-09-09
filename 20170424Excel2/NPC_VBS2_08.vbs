Option Explicit

Dim pnum, ryo

pnum = InputBox("商品番号を入力してください")

'「キャンセル」または数値以外の場合
If pnum = "" Or Not IsNumeric(pnum) Then

	'処理終了
	WScript.Quit

ElseIf pnum >= 101 And pnum <= 105 Then

	ryo = InputBox("購入数を入力してください")

	'「キャンセル」または数値以外の場合
	If ryo = "" Or Not IsNumeric(pnum) Then

		'処理終了
		WScript.Quit 
	
	ElseIf ryo >= 5 Then

		MsgBox "商品番号：" & pnum & vbCr & _
			"購入金額：" & 250 * ryo & "円" 
	Else
		MsgBox "5個以上でご購入ください"
	End If
Else
	MsgBox "その商品は取り扱っていません"

End If
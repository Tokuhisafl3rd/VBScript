Option Explicit

Dim bid1, bid2, bid3

'入札額を入力
bid1 = InputBox("商品Aの入札額を入力してください", "商品A")

'メッセージを表示して「はい」「いいえ」で確認
If MsgBox("入力された金額で入札します", vbYesNo + vbExclamation, _
	"入札確認") = vbNo Then
	MsgBox "処理を中断します"
	WScript.Quit
End If

bid2 = InputBox("商品Bの入札額を入力してください", "商品B")

If MsgBox("入力された金額で入札します", vbYesNo + vbExclamation, _
	"入札確認") = vbNo Then
	MsgBox "処理を中断します"
	WScript.Quit
End If

bid3 = InputBox("商品Cの入札額を入力してください", "商品C")

If MsgBox("入力された金額で入札します", vbYesNo + vbExclamation, _
	"入札確認") = vbNo Then
	MsgBox "処理を中断します"
	WScript.Quit
End If

MsgBox "入札額は" & vbCr & "商品A：" & bid1 & vbCr &"商品B："  _
	& bid2 & vbCr & "商品C：" & bid3 & vbCr & "です",, "入札内容"
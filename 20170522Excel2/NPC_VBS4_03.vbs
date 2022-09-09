Option Explicit

Dim bid1, bid2, bid3

bid1 = InputBox("商品Aの入札額を入力してください", "商品A")

'確認用サブルーチンを引数指定で呼び出す
Call chkBid("商品A", bid1)

bid2 = InputBox("商品Bの入札額を入力してください", "商品B")
Call chkBid("商品B", bid2)

bid3 = InputBox("商品Cの入札額を入力してください", "商品C")
Call chkBid("商品C", bid3)

MsgBox "入札額は" & vbCr & "商品A：" & bid1 & vbCr &"商品B："  _
	& bid2 & vbCr & "商品C：" & bid3 & vbCr & "です",, "入札内容"


'引数付きサブルーチン
Sub chkBid(pName, price)
	If MsgBox("入札額は" & price & "でよろしいですか？", _
		vbYesNo + vbExclamation, pName) = vbNo Then
		MsgBox "処理を中断します"
		WScript.Quit
	End If
End Sub
Option Explicit

Dim ryo

ryo = InputBox("購入数を入力してください")

If ryo < 5 Then

	'少ない場合のメッセージ
	MsgBox "5個以上でご購入ください"
	
ElseIf ryo >= 10 Then

	MsgBox "購入金額：" & 240 * ryo & "円" 
Else

	MsgBox "購入金額：" & 250 * ryo & "円"
End If
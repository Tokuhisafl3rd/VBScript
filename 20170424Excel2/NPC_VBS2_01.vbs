Option Explicit

Dim ryo

'購入数の入力
ryo = InputBox("購入数を入力してください")

'最低購入数のチェック
If ryo < 5 Then MsgBox "5個以上でご購入ください"
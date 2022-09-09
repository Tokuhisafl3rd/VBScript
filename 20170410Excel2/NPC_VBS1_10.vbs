Option Explicit

Dim tanka, ryo

'単価と数量の指定
tanka = InputBox("単価を入力してください")
ryo = InputBox("数量を入力してください")

'税込価格の表示
MsgBox "単価は" & tanka & "円" & vbCr & _
	"数量は" & ryo & "個" & vbCr _
	& "税込価格は" & tanka * ryo * 1.08 & "円です"
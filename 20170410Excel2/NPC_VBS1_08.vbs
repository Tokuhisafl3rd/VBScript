'単価と数量の指定
tanka = 100
ryo = 25

'税込価格の表示
MsgBox "単価は" & tanka & "円" & vbCr & _
	"数量は" & ryo & "個" & vbCr _
	& "税込価格は" & tanka * ryo * 1.08 & "円です"
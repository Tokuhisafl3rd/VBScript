Option Explicit

Dim st, num

st = Timer

'終了条件を指定して繰り返し
Do Until num = 10000000

	num = num + 1

Loop

MsgBox "処理時間は" & Timer - st & "秒です",,"計測結果"
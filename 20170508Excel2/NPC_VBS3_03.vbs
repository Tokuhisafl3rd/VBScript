Option Explicit

Dim st, num

'開始時点を記録
st = Timer

Do
	'カウンター変数を1増やす
	num = num + 1

	'繰り返し回数が1000万になったら処理終了
	If num = 10000000 Then Exit Do

Loop

'処理時間を表示
MsgBox "処理時間は" & Timer - st & "秒です",,"計測結果"
Option Explicit

Dim bDay

'生年月日の入力
bDay = InputBox("生年月日を入力してください")

'経過日数の表示
MsgBox "今日は誕生から" & Date - CDate(bDay) & "日目です"

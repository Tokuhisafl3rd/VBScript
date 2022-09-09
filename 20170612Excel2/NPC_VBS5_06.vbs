Option Explicit

Dim fs, fPath, fTxt

Set fs = CreateObject("Scripting.FileSystemObject")

fPath = fs.GetParentFolderName(WScript.ScriptFullName)

'新しいテキストファイルを作成する
Set fTxt = fs.CreateTextFile(fPath & "\TestData\member06.txt")

'データを書き込む
fTxt.WriteLine "番号：06"
fTxt.WriteLine "川島加奈子"
fTxt.WriteLine "41歳"
fTxt.WriteLine "東京都台東区"
		
fTxt.Close

Set fTxt = Nothing
Set fs = Nothing
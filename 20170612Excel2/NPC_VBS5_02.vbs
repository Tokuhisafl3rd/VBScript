Option Explicit

Dim fs, fPath, fd, fNum

Set fs = CreateObject("Scripting.FileSystemObject")

'このファイルの親フォルダーのパスを取得する
fPath = fs.GetParentFolderName(WScript.ScriptFullName)

'「TestData」フォルダーをFolderオブジェクトとして取得する
Set fd = fs.GetFolder(fPath & "\TestData")

'フォルダー内のファイル数を調べる
fNum = fd.Files.Count

MsgBox fNum & "ファイルです",, "TestData"

Set fd = Nothing
Set fs = Nothing
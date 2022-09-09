Option Explicit

Dim fs

'FileSystemObjectオブジェクトを作成
Set fs = CreateObject("Scripting.FileSystemObject")

'特定のフォルダーの有無を調べる
If fs.FolderExists("C:\Users\○○\Documents\TestData") Then

	MsgBox "指定のフォルダーは存在します",, "存在確認"
Else
	MsgBox "指定のフォルダーは存在しません",, "存在確認"
End If

'使用したオブジェクト変数を解放する
Set fs = Nothing
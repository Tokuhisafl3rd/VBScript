Option Explicit

Dim pw

Do
	'パスワードを入力
	pw = InputBox("パスワードを入力してください","パスワード") 

	'「キャンセル」かどうかを判定
	If IsEmpty(pw) Then
		WScript.Quit

	'解答の判定
	ElseIf pw = "VBScript" Then
		Exit Do
	End If
Loop

MsgBox "パスワードを認証しました",,"認証"
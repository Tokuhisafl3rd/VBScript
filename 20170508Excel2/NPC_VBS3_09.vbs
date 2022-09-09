Option Explicit

Dim num1, num2

num1 = InputBox("”»’è‚µ‚½‚¢”’l‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢", "‘f””»’è")

If IsEmpty(num1) Then
	WScript.Quit

ElseIf IsNumeric(num1) Then

	num1 = Int(num1)

	If num1 < 2 Then
		MsgBox num1 & "‚Í‘f”‚Å‚Í‚ ‚è‚Ü‚¹‚ñ",, "‘f””»’è"
		WScript.Quit
	End If
Else
	MsgBox "”’l‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢",, "‘f””»’è"
	WScript.Quit
End If

'2‚©‚ç”’l‚Ì³‚Ì•½•ûª‚Ì’l‚Ü‚ÅŒJ‚è•Ô‚µ
For num2 = 2 To Int(Sqr(num1))

	'Š„‚èØ‚ê‚é‚©‚Ç‚¤‚©‚ğ”»’è
	If num1 Mod num2 = 0 Then
		MsgBox num1 & "‚Í‘f”‚Å‚Í‚ ‚è‚Ü‚¹‚ñ",, "‘f””»’è"
		WScript.Quit
	End If
Next

MsgBox num1 & "‚Í‘f”‚Å‚·",, "‘f””»’è"

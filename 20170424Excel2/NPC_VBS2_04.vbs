Option Explicit

Dim ryo

ryo = InputBox("w“ü”‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢")

If ryo < 5 Then

	'­‚È‚¢ê‡‚ÌƒƒbƒZ[ƒW
	MsgBox "5ŒÂˆÈã‚Å‚²w“ü‚­‚¾‚³‚¢"
	
ElseIf ryo >= 10 Then

	MsgBox "w“ü‹àŠzF" & 240 * ryo & "‰~" 
Else

	MsgBox "w“ü‹àŠzF" & 250 * ryo & "‰~"
End If
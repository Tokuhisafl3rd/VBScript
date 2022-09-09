Option Explicit

Dim pnum, ryo

pnum = InputBox("¤•i”Ô†‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢")

If pnum >= 101 And pnum <= 105 Then

	'w“ü”‚Ì“ü—Í
	ryo = InputBox("w“ü”‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢")

	'w“ü”‚Ì”»’è
	If ryo >= 5 Then

		MsgBox "¤•i”Ô†F" & pnum & vbCr & _
			"w“ü‹àŠzF" & 250 * ryo & "‰~" 
	Else
		MsgBox "5ŒÂˆÈã‚Å‚²w“ü‚­‚¾‚³‚¢"
	End If
Else
	MsgBox "‚»‚Ì¤•i‚Íæ‚èˆµ‚Á‚Ä‚¢‚Ü‚¹‚ñ"

End If
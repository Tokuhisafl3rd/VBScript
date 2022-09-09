Option Explicit

Dim num1, num2, num3

num1 = InputBox("0`15‚Ì®”‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢", "’l‚Ì“ü—Í")

num2 = 1

For num3 = 2 To num1
	num2 = num2 * num3
Next

MsgBox num1 & "‚ÌŠKæF" & num2, , "ŠKæ‚Ì’l"

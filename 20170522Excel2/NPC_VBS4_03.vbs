Option Explicit

Dim bid1, bid2, bid3

bid1 = InputBox("���iA�̓��D�z����͂��Ă�������", "���iA")

'�m�F�p�T�u���[�`���������w��ŌĂяo��
Call chkBid("���iA", bid1)

bid2 = InputBox("���iB�̓��D�z����͂��Ă�������", "���iB")
Call chkBid("���iB", bid2)

bid3 = InputBox("���iC�̓��D�z����͂��Ă�������", "���iC")
Call chkBid("���iC", bid3)

MsgBox "���D�z��" & vbCr & "���iA�F" & bid1 & vbCr &"���iB�F"  _
	& bid2 & vbCr & "���iC�F" & bid3 & vbCr & "�ł�",, "���D���e"


'�����t���T�u���[�`��
Sub chkBid(pName, price)
	If MsgBox("���D�z��" & price & "�ł�낵���ł����H", _
		vbYesNo + vbExclamation, pName) = vbNo Then
		MsgBox "�����𒆒f���܂�"
		WScript.Quit
	End If
End Sub
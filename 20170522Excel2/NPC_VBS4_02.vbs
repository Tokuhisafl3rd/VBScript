Option Explicit

Dim bid1, bid2, bid3

bid1 = InputBox("���iA�̓��D�z����͂��Ă�������", "���iA")

'�m�F�p�T�u���[�`�����Ăяo��
Call chkBid

bid2 = InputBox("���iB�̓��D�z����͂��Ă�������", "���iB")
Call chkBid

bid3 = InputBox("���iC�̓��D�z����͂��Ă�������", "���iC")
Call chkBid

MsgBox "���D�z��" & vbCr & "���iA�F" & bid1 & vbCr &"���iB�F"  _
	& bid2 & vbCr & "���iC�F" & bid3 & vbCr & "�ł�",, "���D���e"


'�m�F�������T�u���[�`����
Sub chkBid()
	If MsgBox("���͂��ꂽ���z�œ��D���܂�", vbYesNo + vbExclamation, _
		"���D�m�F") = vbNo Then
		MsgBox "�����𒆒f���܂�"
		WScript.Quit
	End If
End Sub
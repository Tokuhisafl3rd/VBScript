Option Explicit

Dim bid1, bid2, bid3

'���D�z�����
bid1 = InputBox("���iA�̓��D�z����͂��Ă�������", "���iA")

'���b�Z�[�W��\�����āu�͂��v�u�������v�Ŋm�F
If MsgBox("���͂��ꂽ���z�œ��D���܂�", vbYesNo + vbExclamation, _
	"���D�m�F") = vbNo Then
	MsgBox "�����𒆒f���܂�"
	WScript.Quit
End If

bid2 = InputBox("���iB�̓��D�z����͂��Ă�������", "���iB")

If MsgBox("���͂��ꂽ���z�œ��D���܂�", vbYesNo + vbExclamation, _
	"���D�m�F") = vbNo Then
	MsgBox "�����𒆒f���܂�"
	WScript.Quit
End If

bid3 = InputBox("���iC�̓��D�z����͂��Ă�������", "���iC")

If MsgBox("���͂��ꂽ���z�œ��D���܂�", vbYesNo + vbExclamation, _
	"���D�m�F") = vbNo Then
	MsgBox "�����𒆒f���܂�"
	WScript.Quit
End If

MsgBox "���D�z��" & vbCr & "���iA�F" & bid1 & vbCr &"���iB�F"  _
	& bid2 & vbCr & "���iC�F" & bid3 & vbCr & "�ł�",, "���D���e"
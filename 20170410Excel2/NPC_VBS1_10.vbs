Option Explicit

Dim tanka, ryo

'�P���Ɛ��ʂ̎w��
tanka = InputBox("�P������͂��Ă�������")
ryo = InputBox("���ʂ���͂��Ă�������")

'�ō����i�̕\��
MsgBox "�P����" & tanka & "�~" & vbCr & _
	"���ʂ�" & ryo & "��" & vbCr _
	& "�ō����i��" & tanka * ryo * 1.08 & "�~�ł�"
Option Explicit

Dim bDay

'���N�����̓���
bDay = InputBox("���N��������͂��Ă�������")

'�o�ߓ����̕\��
MsgBox "�����͒a������" & Date - CDate(bDay) & "���ڂł�"

Option Explicit

Dim st, num

'�J�n���_���L�^
st = Timer

Do
	'�J�E���^�[�ϐ���1���₷
	num = num + 1

	'�J��Ԃ��񐔂�1000���ɂȂ����珈���I��
	If num = 10000000 Then Exit Do

Loop

'�������Ԃ�\��
MsgBox "�������Ԃ�" & Timer - st & "�b�ł�",,"�v������"
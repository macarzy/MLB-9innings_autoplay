[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=dfb33df9-2c96-4a6d-bddb-536242c0fd78
Description=�����ǿ�X�Ʀr�ǦC��
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]
.
�i�\�໡���j.
�����ǿ�X�Ʀr�A�i�H�]�w�}�l���Ʀr�M�������Ʀr�C.
.
�i�@�̤��Сj����M�a�� �L�~�]�� (godeden).
.
.

[Script]
// ����������������������������������������������������������������������
// �i�\�໡���j�Ʀr�ǦC����X�i�ҪO�j
// �i�ϥλ����j�b�u�X����ܮؤ���J�_�l���X(�̤p�Ʀr)�򵲧����X(�̤j�Ʀr)
// �i���ܻ����j�Цb�O�ƥ�������
// ���������������������������������w�ܶq��������������������������������
dim a,b
// ���������������������������������D�{�ǡ�������������������������������
// �i�}���}�l�j
a=InputBox("��J�_�l���X")
b=InputBox("��J�������X")
// �i�o��cint(x)�O���ܶq���ഫ���Ʀr�j
a=cint(a)
b=cint(b)
While a<=b
    SayString a
    KeyPress 32,1
    a=a+1
Wend
// ���������������������������������ơ��`��������������������������������
// �i�}���g�L�ק�j
// �i�Ч@����j2005
// �i�Ʀr�O���j�������������������Q
// ����������������������������������������������������������������������

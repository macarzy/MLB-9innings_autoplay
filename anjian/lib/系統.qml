[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//



[General]
MacroID=a2f21d32-ba73-4078-ab71-8b1c505273bd
SyntaxVersion=2
Description=�t��

[Script]
//�Цb�U���g�W�z���l�{�ǩΨ��
//�g���O�s��A�b���@�R�O�w�W�I���k��ÿ�ܡu��s�v�Y�i


Sub �����i�{(�M���W��)
    //Call Lib.�t��.�����i�{("notepad.exe")
    Dim strComputer, objWMIService, colProcessList, objProcess
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & �M���W�� & "'")
    For Each objProcess in colProcessList
        objProcess.Terminate
    Next
End Sub


//�s�@�G�@����
//����G2010.10.07
//�ק�G2010.10.07
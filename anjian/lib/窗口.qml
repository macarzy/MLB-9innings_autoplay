[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//


[General]
MacroID=45322cb5-c49c-4570-97d6-48c53e5170ed

[Script]
Function �o�칫�Цb���f�W��m()
    //�Ҥl�GMsgBox lib.���f.�o�칫�Цb���f�W��m()
    Dim ���ФU�y�`,x,y,����
    ���ФU�y�` = Plugin.Window.MousePoint()
    Call GetCursorPos(x,y)
    ���� = Plugin.Window.GetClientRect(���ФU�y�`)
    Dim MyArray
    MyArray = Split(����, "|")
    �o�칫�Цb���f�W��m = x - Clng(MyArray(0)) & "|" & y - Clng(MyArray(1)) 
End Function

Function �u�X��ܮ�(���ܤ��e,���ݮɶ�,���ܼ��D,��ܼ˦�)
    //�Ҥl�GMsgBox "�A��ܪ��O�G" & lib.���f.�u�X��ܮ�("���ܤ��e",0,"���ܼ��D",68)
    //�ԲӨϥΰѦҳo�̡Ghttp://bbs.vrbrothers.com/viewthread.php?tid=7662
    Dim obj
    Set obj = CreateObject("WScript.Shell")
    �u�X��ܮ�=Cint(obj.Popup(���ܤ��e,���ݮɶ�,���ܼ��D,��ܼ˦�))
    Set obj = Nothing
End Function





//�s�@�G�@����
//����G2009.12.22
//�ק�G2010.01.19



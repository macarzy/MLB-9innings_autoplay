[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//



[General]
MacroID=d5e279c7-3d49-44d8-ba83-c1700ece0c96

[Script]
Function �d��̹��Ϥ��ƶq(������,�W����,�k����,�U����,�Ϥ����|,�ۦ���)
    //�Ҥl�GMsgBox lib.�Ϲ�.�d��̹��Ϥ��ƶq(0,0,800,300,"C:\�ϼ�.bmp",0.9)
    //A1.B1.C1.D1  �O���F�K��]�m��Ϫ��d��
    Dim A1,B1,C1,D1,a,b,c,d,n,x,y,H
    A1=������
    B1=�W����
    C1=�k����
    D1=�U����
    //(a.b.c.d)���n�ק�
    a=A1: b=B1: c=C1: d=D1
    //n�O�Ϥ����ƶq
    n=0
    Rem �`���j��
    Call FindPic(a,b,c,d,�Ϥ����|,�ۦ���,x,y)
    If (x>=0 and y>=0 and y=b and a=A1) Or (x>=0 and y>=0 and y=b and a<>A1) Or (x>=0 and y>=0 and a=A1 and y<>b) Then
        n=n+1: H=y:  a=x+1: b=y
        Goto �`���j��
    ElseIf a>A1 Then
        a=A1: b=H+1
        Goto �`���j��
    End If 
    �d��̹��Ϥ��ƶq = n
End Function
Sub �L���̹��I��(������,�W����,�k����,�U����,�O�s�Ϥ����|,���榡)
    //�Ҥl�GCall lib.�Ϲ�.�L���̹��I��(0,0,100,100,"C:\","bmp")
    Dim �ɶ�
    �ɶ�=Year(Now)& Month(Now)& Day(Now)& Hour(Now)& Minute(Now)& Second(Now)
    Call Plugin.Pic.PrintScreen(������,�W����,�k����,�U����,�O�s�Ϥ����| & �ɶ� & "." & ���榡)
End Sub



//�s�@�G�@����
//����G2009.12.22




[General]
SyntaxVersion=2
MacroID=703c757b-2ea1-46a1-8191-3111913e3930
[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//



[Script]
//API�եΡ]���ա^�O�@�ӫܱj�j���\��A�i�H��K���ϥΨt�Φ۱a��API�W�j�}���\��C
//���P��API�եΤS�O�@�ӫ����Ϊ��\��A�u�����xAPI���@�̡]�@��O�{�ǭ��^�A�~�঳��O�Φn�C
//�]���A������F��API�եΥ\�઺�ϥκA�����ӬO�G
//1�B���y�NAPI�եΥ\��]�ˬ��w�R�O�A�o�G���j�a�ϥΡC
//2�B�Ϲ����Q��API�եΥ\��}�o�g���s�B�ק�ƾګʥ]�B�ק�C���Ȥ�ݼƾڵ��i��I�ǲĤT�誾�Ѳ��v���\��C
//3�B���@�ǯS��API�A�ȮɵL�k������A�i�H�δ�����������API�������C

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Function �d�䵡�f�y�`(���f���W,���f���D)
    //�Ҥl�GMsgBox lib.API.�d�䵡�f�y�`("notepad",0)
    Dim sClass, sTitle
    If Cstr(���f���W) = "0" Then
        �d�䵡�f�y�` = FindWindow(vbNullString, Cstr(���f���D))
    ElseIf Cstr(���f���D) = "0" Then
        �d�䵡�f�y�` = FindWindow(Cstr(���f���W), vbNullString)
    Else
        �d�䵡�f�y�` = FindWindow(Cstr(���f���W), Cstr(���f���D))
    End If
End Function
Sub �E�����f�øm�e(���f�y�`)
    //���f�y�` = lib.API.�d�䵡�f�y�`("notepad",0)
    //�Ҥl�GCall lib.API.�E�����f�øm�e(���f�y�`)
    Dim SW_SHOWNORMAL
    SW_SHOWNORMAL = 1
    If ���f�y�` <> 0 Then 
        Call ShowWindow(���f�y�`, SW_SHOWNORMAL)
        Call SetForegroundWindow(���f�y�`)
    End If
End Sub
Sub �]�m���f�z����(���f�y�`, �z����)
    //�z���סG0~255
    //���f�y�` = lib.API.�d�䵡�f�y�`("notepad",0)
    //�Ҥl�GCall lib.API.�]�m���f�z����(���f�y�`,100)
    Dim GWL_EXSTYLE, LWA_ALPHA, WS_EX_LAYERED 
    GWL_EXSTYLE = (-20)
    LWA_ALPHA = &H2
    WS_EX_LAYERED = &H80000
    Dim Rt
    If ���f�y�` <> 0 And �z����>=0 And �z����<=255 Then
        Rt = GetWindowLong(���f�y�`, GWL_EXSTYLE)
        Rt = Rt Or WS_EX_LAYERED
        Call SetWindowLong(���f�y�`, GWL_EXSTYLE, Rt)
        Call SetLayeredWindowAttributes(���f�y�`, 0, �z����, LWA_ALPHA)
    End If
End Sub
Sub �]�m���f���Ь�z(���f�y�`)
    //���f�y�` = lib.API.�d�䵡�f�y�`("notepad",0)
    //�Ҥl�GCall lib.API.�]�m���f���Ь�z(���f�y�`)
    Dim GWL_EXSTYLE, WS_EX_TRANSPARENT, WS_EX_LAYERED 
    GWL_EXSTYLE = (-20)
    WS_EX_TRANSPARENT = &H20
    WS_EX_LAYERED = &H80000
    Dim rtn
    If ���f�y�` <> 0 Then
        rtn = GetWindowLong(���f�y�`, GWL_EXSTYLE)
        rtn = rtn Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
        SetWindowLong ���f�y�`, GWL_EXSTYLE, rtn
    End If
End Sub
Sub �������f�]�m(���f�y�`, �Ҧ�)
    //�Ҧ��G0(�@�뵡��)~1(UpdateLayeredWindow�e����)
    //���f�y�` = lib.API.�d�䵡�f�y�`("notepad",0)
    //�Ҥl�GCall lib.API.�������f�]�m(���f�y�`,0)
    Dim GWL_EXSTYLE, WS_EX_LAYERED 
    GWL_EXSTYLE = (-20) 
    WS_EX_LAYERED = &H80000
    If ���f�y�` <> 0 Then
        If �Ҧ� = 0 then
            '�������Ь�z(�@�뵡��)  
            SetWindowLong ���f�y�`, GWL_EXSTYLE, 0
        Else 
            '�������Ь�z(UpdateLayeredWindow�e����)  
            SetWindowLong ���f�y�`, GWL_EXSTYLE, WS_EX_LAYERED
        End If
    End If
End Sub
Sub �B��{��(�{�Ǹ��|)
    //�Ҥl�GCall Lib.API.�B��{��("E:\�����C��\S�T�굦6\GLSGC.exe")
    Dim P, I, DirPath, ExeName
    P = Split(�{�Ǹ��|, "\")
    For I = 0 To UBound(P) - 1
        DirPath = DirPath & P(I) & "\"
    Next
    ExeName = P(UBound(P))
    ShellExecute GetDesktopWindow, "open", ExeName, vbNullString, DirPath, 5
End Sub







//�s�@�G�@����
//����G2010.11.10
//�ק�G2011.06.18

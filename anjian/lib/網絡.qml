[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//



[General]
MacroID=a1c52b4e-dc44-4ded-887c-ff94144acf68

[Script]
Function ��o���������(�����a�})
    //�����G������{����奻���e�A�p�GMsgBox lib.����.��o���������("http://www.jdyou.com/test.txt")
    //�Ҥl�GMsgBox lib.����.��o���������("http://www.jdyou.com")
    Dim xmlHttp, xmlBody, xmlUrl
    Dim ThisCharCode ,NextCharCode ,BytesToBstr
    If InStr(�����a�}, "http://") = 0 Then 
        xmlUrl = "http://" & �����a�}
    Else
        xmlUrl = �����a�}
    End if
    Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
    xmlHttp.Open "Get", xmlUrl, False
    xmlHttp.Send
    xmlBody = xmlHttp.ResponseBody
    Set xmlHttp = Nothing  
    ��o��������� = ""
    If Len(xmlBody) = 0 Then Exit Function
    Set ObjStream = CreateObject("Adodb.Stream")
    With ObjStream
        .Type = 1
        .Mode = 3
        .Open
        .Write xmlBody
        .Position = 0
        .Type = 2
        .Charset = "GB2312"
        BytesToBstr = .ReadText
        .Close
    End With
    Set ObjStream = Nothing    
    ��o��������� = BytesToBstr
End Function
Function ��o�~��IP�a�}()   
    //�Ҥl�GMsgBox lib.����.��o�~��IP�a�}()
    Dim �������e,�}�l��m,������m
    �������e = lib.����.��o���������("http://iframe.ip138.com/ic.asp")
    �}�l��m = inStr(�������e,"[") + 1
    ������m = inStr(�������e,"]") - �}�l��m
    ��o�~��IP�a�} = Mid(�������e,�}�l��m,������m)
End Function
Function �o�e�l��(�A���l�c�b��, �A���l�c�K�X, �o�e�l��a�}, �l��D�D, �l�󤺮e, �l�����) 
    //�Ҥl�GMsgBox lib.����.�o�e�l��("ceshi0000001@163.com","ceshi000001","ceshi0000001@163.com","�l��D�D","�l�󤺮e","")
    Dim You_ID,MS_Space,Email
    '�b���M�A�Ⱦ����� 
    You_ID = Split(�A���l�c�b��, "@") 
    '�o�ӬO�����n���A���L�i�H��ߪ��ơA���|�q�L�L�n�o�e�l�� 
    MS_Space = "http://schemas.microsoft.com/cdo/configuration/" 
    Set Email = CreateObject("CDO.Message") 
    '�o�Ӥ@�w�n�M�o�e�l�󪺱b���@��
    Email.From = �A���l�c�b�� 
    'Execute "Email.to = �o�e�l��a�}"
    Email.CC = �o�e�l��a�}
    Email.Subject = �l��D�D
    Email.Textbody = �l�󤺮e 
    If �l����� <> "" Then 
        Email.AddAttachment �l����� 
    End If 
    With Email.Configuration.Fields 
        '�o�H�ݤf 
        .Item(MS_Space & "sendusing") = 2 
        'SMTP�A�Ⱦ��a�} 
        .Item(MS_Space & "smtpserver") = "smtp." & You_ID(1) 
        'SMTP�A�Ⱦ��ݤf 
        .Item(MS_Space & "smtpserverport") = 25 
        .Item(MS_Space & "smtpauthenticate") = 1
        .Item(MS_Space & "sendusername") = You_ID(0) 
        .Item(MS_Space & "sendpassword") = �A���l�c�K�X  
        .Update 
    End With 
    '�o�e�l�� 
    Email.Send 
    '�����ե� 
    Set Email = Nothing 
    �o�e�l�� = True
    '�p�G�S��������~�H���A�h��ܵo�e���\,�_�h�o�e���� 
    If Err Then 
        Err.Clear 
        �o�e�l�� = False 
    End If 
End Function 
Function ��������ɶ�()
    //�Ҥl�GMsgBox "��e�зǮɶ����G" & lib.����.��������ɶ�()
    //�P�_�GIf NowTime>CDate("2010-5-9") Then
    Dim SvrName(7),xPost,HttpAdd,NowTime,StartTime,i
    StartTime=Now 
    //SvrName(0) = "time-a.nist.gov"
    SvrName(1) = "time-a.timefreq.bldrdoc.gov"
    SvrName(2) = "time-b.timefreq.bldrdoc.gov"
    SvrName(3) = "time-c.timefreq.bldrdoc.gov"
    SvrName(4) = "utcnist.colorado.edu"
    SvrName(5) = "time.nist.gov"
    SvrName(6) = "nist1.datum.com"
    SvrName(7) = "nist1.aol-ca.truetime.com"
    Set xPost=createObject("Microsoft.XMLHTTP") 
    NowTime=""
    Do While NowTime=""
        For i=1 to 7
            NowTime=""
            HttpAdd="Http://" & SvrName(i) & ":13"
            xPost.Open "Put", HttpAdd, False
            xPost.Send
            Delay 10
            If xPost.readyState=4 Then
                NowTime=mid(xPost.responsetext, 8, 17)
                If NowTime<>"" Then
                    NowTime=CDate(NowTime) + 8 / 24
                    Exit Do
                Else
                    xPost.abort
                    NowTime=""
                End If
            End If
        Next
        If DateDiff("s", StartTime, Now)>=30 And NowTime="" Then
            Msgbox "�нT�w�A�w�g�s���W�F���p���I", 0, "��������ɶ�"
            Exit Do 
        End If
    Loop
    xPost.abort
    Set xPost=Nothing
    ��������ɶ�=NowTime
End Function






//�s�@�G�@����
//����G2009.12.30
//�ק�G2011.04.19


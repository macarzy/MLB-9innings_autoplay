[General]
SyntaxVersion=2
MacroID=1f3dfae4-1654-4de0-a32c-37e85ed16fd8
[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//




[Script]
Function Ū�����w��奻���e(�奻���|, ���)
    //�Ҥl�GMsgbox lib.���.Ū�����w��奻���e("C:\Log.txt", 3)
    Dim fso,myfile,i,flag,tempp 
    flag = 1
    Set fso=CreateObject("scripting.FileSystemObject")
    If fso.FileExists(�奻���|) then
        Set myfile = fso.openTextFile(�奻���|,1,false)
    Else
        flag = 0
    End If
    For i=1 To ��� - 1
        If Not myfile.AtEndOfLine Then
            myfile.SkipLine
            tempp = myfile.Line
        End If
    Next
    If flag = 1 Then
        If Not myfile.AtEndOfLine Then
            Ū�����w��奻���e = myfile.ReadLine
        Else
            Ū�����w��奻���e = "���X�I"
        End If
        myfile.close
    Else
        Ū�����w��奻���e = "��󤣦s�b�I"
    End If
    Set fso = Nothing
End Function
Sub �R�����w��奻���e(�奻���|, ���)
    //�Ҥl�GCall lib.���.�R�����w��奻���e("C:\log.txt",5)
    Dim ForReading ,ForWriting 
    ForReading = 1
    ForWriting = 2
    Dim objFSO,objFile,strLine,strNewFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(�奻���|,ForReading)
    Do Until objFile.AtEndOfStream
        strLine = objFile.Readline       
        If ��� = objFile.Line-1 or ��� = 0 then 
            strNewFile = strNewFile 
        Else
            strNewFile = strNewFile & strLine & vbcrlf
        End If
    Loop
    objFile.Close
    Set objFile = objFSO.OpenTextFile(�奻���|,ForWriting)
    objFile.Write strNewFile
    objFile.Close
    Set objFSO = Nothing
End Sub
Sub �������w��奻���e(�奻���|, �奻���e, ���)
    //�Ҥl�GCall lib.���.�������w��奻���e("C:\log.txt","�奻���e",5)
    Dim ForReading ,ForWriting 
    ForReading = 1
    ForWriting = 2
    Dim objFSO,objFile,strLine,strNewFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(�奻���|,ForReading)
    Do Until objFile.AtEndOfStream
        strLine = objFile.Readline  
        If ��� = objFile.Line-1 or ��� = 0 then
            strNewFile = strNewFile & �奻���e & vbcrlf
        Else
            strNewFile = strNewFile & strLine & vbcrlf
        End If
    Loop
    objFile.Close
    Set objFile = objFSO.OpenTextFile(�奻���|,ForWriting)
    objFile.Write strNewFile
    objFile.Close
    Set objFSO = Nothing
End Sub
Sub ���J�奻���e����w��(�奻���|, �奻���e, ���)
    //�Ҥl�GCall lib.���.���J�奻���e����w��("C:\log.txt","�奻���e",5)
    Dim ForReading ,ForWriting 
    ForReading = 1
    ForWriting = 2
    Dim objFSO,objFile,strLine,strNewFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(�奻���|,ForReading)
    Do Until objFile.AtEndOfStream
        strLine = objFile.Readline       
        strNewFile = strNewFile & strLine & vbcrlf
        If ��� = objFile.Line-1 or ��� = 0 then
            strNewFile = strNewFile & �奻���e & vbcrlf
        End If
    Loop
    objFile.Close
    Set objFile = objFSO.OpenTextFile(�奻���|,ForWriting)
    objFile.Write strNewFile
    objFile.Close
    Set objFSO = Nothing
End Sub
Function �M�����w�ؿ��U�Ҧ����W(��󧨸��|)
    //�`�N�G��^���O�Ʋ��ܶq�A�s�x�ۨC�@�Ӥ��W�C
    //�Ҥl�G�Ʋ� = lib.���.�M�����w�ؿ��U�Ҧ����W("C:\")
    //      For i=0 to UBound(�Ʋ�)-1
    //          TracePrint �Ʋ�(i)
    //      Next
    Dim ���W,fso,folder,f,files
    ���W = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.getfolder(��󧨸��|)
    Set files = folder.files
    For Each f In files
        ���W = ���W & f.name & ","
    Next
    Set fso = Nothing
    //�M�����w�ؿ��U�Ҧ����W = ���W
    �M�����w�ؿ��U�Ҧ����W = Split(���W, ",")
End Function
Function �M�����w�ؿ��U�Ҧ���󧨦W(��󧨸��|)
    //�`�N�G��^���O�Ʋ��ܶq�A�s�x�ۨC�@�Ӥ�󧨦W�C
    //�Ҥl�G�Ʋ� = lib.���.�M�����w�ؿ��U�Ҧ���󧨦W("C:\")
    //      For i=0 to UBound(�Ʋ�)-1
    //          TracePrint �Ʋ�(i)
    //      Next
    Dim ��󧨦W,fso,folder,f,files
    ��󧨦W = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.getfolder(��󧨸��|)
    Set files = folder.SubFolders
    For Each f In files
        ��󧨦W = ��󧨦W & f.name & ","
    Next
    Set fso = Nothing
    //�M�����w�ؿ��U�Ҧ���󧨦W = ��󧨦W
    �M�����w�ؿ��U�Ҧ���󧨦W = Split(��󧨦W, ",")
End Function
Function �P�_��󧨬O�_�s�b(��󧨸��|)
    //�Ҥl�GMsgbox lib.���.�P�_��󧨬O�_�s�b("c:\WINDOWS")
    Dim fso 
    Set fso = CreateObject("Scripting.FileSystemObject")
    �P�_��󧨬O�_�s�b = fso.FolderExists(��󧨸��|)
    Set fso = Nothing
End Function



//�s�@�G�@����
//����G2009.12.22
//�ק�G2011.05.03



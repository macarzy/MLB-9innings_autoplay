[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//


[General]
MacroID=3626b374-83c9-406b-978f-7ef9fe123bc1



[Script]
Sub ��L�զX��(��L�X,�����覡)
    //�Ҥl�GCall lib.��L.��L�զX��("Ctrl + Alt + A",0)
    //�����覡�G�i0���q�����A1�w������A2�W�ż����j
    //��h�i����X�j�i�H�ۦ�K�[
    Dim ������,���U��,�\����,��V��,�r����,�Ʀr��,�Ÿ���
    ������ = "CTRL 17,ALT 18,SHIFT 16,LCTRL 162,LALT 164,LSHIFT 160,RCTRL 163,RALT 165,RSHIFT 161,WIN 91"
    ���U�� = "CTRL 17,ALT 18,SHIFT 16,LCTRL 162,LALT 164,LSHIFT 160,RCTRL 163,RALT 165,RSHIFT 161"
    ��V�� = "DOWN 40,UP 38,LEFT 37,RIGHT 39"
    �\���� = "F1 112,F2 113,F3 114,F4 115,F5 116,F6 117,F7 118,F8 119,F9 120,F10 121,F11 122,F12 123,HOME 36,END 35,PAGEDOWN 34,PAGEUP 33,ESC 27,ENTER 13,SPACE 32"
    �r���� = "A 65,B 66,C 67,D 68,E 69,F 70,G 71,H 72,I 73,J 74,K 75,L 76,M 77,N 78,O 79,P 80,Q 81,R 82,S 83,T 84,U 85,V 86,W 87,X 88,Y 89,Z 90"
    �Ʀr�� = "0 48,1 49,2 50,3 51,4 52,5 53,6 54,7 55,8 56,9 57"
    �Ÿ��� = "~ 192,` 192,- 189,= 187,[ 219,] 221,\ 220,/ 191,? 191,< 188,> 190"
    //�w���˴�1
    Dim �ন�j�g,�h���Ů�,i
    �ন�j�g = UCase(��L�X)
    �h���Ů� = Replace(�ন�j�g," ","")
    Dim ���Υ[��,�[���ƶq
    ���Υ[�� = Split(�h���Ů�,"+") 
    �[���ƶq = UBound(���Υ[��)
    If �[���ƶq>0 And �[���ƶq<3 Then        
        If InStr(������,���Υ[��(0))>0 And ���Υ[��(0)<>"" Then 
            //�p�ⱱ����X
            Dim ��,����
            �� = Split(������,",") 
            For i=0 To UBound(��)
                If InStr(��(i),���Υ[��(0))>0 Then
                    ���� = Split(��(i)," ")   
                    Exit For 
                End If 
            Next       
            Dim ��,���U,����
            If �[���ƶq = 1 Then   
                If ���Υ[��(1)<>���Υ[��(0) And ���Υ[��(1)<>"" Then    
                    Dim �X�k1(4)
                    �X�k1(0) = InStr(�\����,���Υ[��(1))
                    �X�k1(1) = InStr(��V��,���Υ[��(1))
                    �X�k1(2) = InStr(�r����,���Υ[��(1))
                    �X�k1(3) = InStr(�Ʀr��,���Υ[��(1))
                    �X�k1(4) = InStr(�Ÿ���,���Υ[��(1))
                    //�w���˴�2
                    If �X�k1(0)>0 Or �X�k1(1)>0 Or �X�k1(2)>0 Or �X�k1(3)>0 Or �X�k1(4)>0 Then  
                        //�p�������X
                        �� = Split(�r����,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(1) & " ")>0 Then
                                ���U = Split(��(i)," ")  
                                Goto ����1 
                            End If 
                        Next  
                        �� = Split(�Ʀr��,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(1) & " ")>0 Then
                                ���U = Split(��(i)," ") 
                                Goto ����1 
                            End If 
                        Next         
                        �� = Split(��V��,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(1) & " ")>0 Then
                                ���U = Split(��(i)," ")   
                                Goto ����1 
                            End If 
                        Next  
                        �� = Split(�\����,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(1) & " ")>0 Then
                                ���U = Split(��(i)," ")
                                Goto ����1 
                            End If 
                        Next 
                        �� = Split(�Ÿ���,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(1) & " ")>0 Then
                                ���U = Split(��(i)," ")  
                                Goto ����1 
                            End If 
                        Next             
                        Rem ����1 
                        //����L�զX��
                        If �����覡 = 0 Then
                            KeyDown Clng(����(1)), 1
                            KeyPress Clng(���U(1)), 1
                            KeyUp Clng(����(1)), 1
                        ElseIf �����覡 = 1 Then
                            KeyDownH Clng(����(1)), 1
                            KeyPressH Clng(���U(1)), 1
                            KeyUpH Clng(����(1)), 1
                        ElseIf �����覡 = 2 Then 
                            KeyDownS Clng(����(1)), 1
                            KeyPressS Clng(���U(1)), 1
                            KeyUpS Clng(����(1)), 1
                        End If 
                        Exit Sub
                    End If 
                End If 
            ElseIf �[���ƶq = 2 Then
                If ���Υ[��(2)<>���Υ[��(1) And ���Υ[��(2)<>"" Then 
                    Dim �X�k2(5)
                    �X�k2(0) = InStr(���U��,���Υ[��(2))
                    �X�k2(1) = InStr(�\����,���Υ[��(2))
                    �X�k2(2) = InStr(��V��,���Υ[��(2))
                    �X�k2(3) = InStr(�r����,���Υ[��(2))
                    �X�k2(4) = InStr(�Ʀr��,���Υ[��(2))
                    �X�k2(5) = InStr(�Ÿ���,���Υ[��(2))
                    //�w���˴�3
                    If �X�k2(0)>0 Or �X�k2(1)>0 Or �X�k2(2)>0 Or �X�k2(3)>0 Or �X�k2(4)>0 Or �X�k2(5)>0 Then
                        //�p�������X
                        �� = Split(�r����,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(2) & " ")>0 Then
                                ���� = Split(��(i)," ")  
                                Goto ����2 
                            End If 
                        Next  
                        �� = Split(�Ʀr��,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(2) & " ")>0 Then
                                ���� = Split(��(i)," ") 
                                Goto ����2 
                            End If 
                        Next 
                        �� = Split(��V��,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(2) & " ")>0 Then
                                ���� = Split(��(i)," ")   
                                Goto ����2 
                            End If 
                        Next    
                        �� = Split(�\����,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(2) & " ")>0 Then
                                ���� = Split(��(i)," ")
                                Goto ����2 
                            End If 
                        Next  
                        �� = Split(�Ÿ���,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(2) & " ")>0 Then
                                ���� = Split(��(i)," ")  
                                Goto ����2 
                            End If 
                        Next 
                        Rem ����2
                        //�p�⻲�U��X
                        �� = Split(���U��,",") 
                        For i=0 To UBound(��)
                            If InStr(��(i),���Υ[��(1) & " ")>0 Then
                                ���U = Split(��(i)," ")
                                Exit For 
                            End If 
                        Next  
                        //����L�զX��
                        If �����覡 = 0 Then
                            KeyDown Clng(����(1)), 1
                            KeyDown Clng(���U(1)), 1
                            KeyPress Clng(����(1)), 1
                            KeyUp Clng(���U(1)), 1
                            KeyUp Clng(����(1)), 1
                        ElseIf �����覡 = 1 Then
                            KeyDownH Clng(����(1)), 1
                            KeyDownH Clng(���U(1)), 1
                            KeyPressH Clng(����(1)), 1
                            KeyUpH Clng(���U(1)), 1
                            KeyUpH Clng(����(1)), 1
                        ElseIf �����覡 = 2 Then
                            KeyDownS Clng(����(1)), 1
                            KeyDownS Clng(���U(1)), 1
                            KeyPressS Clng(����(1)), 1
                            KeyUpS Clng(���U(1)), 1
                            KeyUpS Clng(����(1)), 1
                        End If 
                        Exit Sub
                    End If 
                End If 
            End If
            //�q�L�w��            
        End If
    End If 
End Sub


Sub ��L�����(��X��,�����覡,�@����)
    //�Ҥl�GCall lib.��L.��L�����("A,B,C,SPACE,D,E,F,G",0,50)
    //�����覡�G�i0���q�����A1�w������A2�W�ż����j
    //��h�i����X�j�i�H�ۦ�K�[
    Dim ������,���U��,�\����,��V��,�r����,�Ʀr��,�Ÿ���,�զX��
    ������ = "CTRL 17,ALT 18,SHIFT 16,LCTRL 162,LALT 164,LSHIFT 160,RCTRL 163,RALT 165,RSHIFT 161,WIN 91"
    �\���� = "F1 112,F2 113,F3 114,F4 115,F5 116,F6 117,F7 118,F8 119,F9 120,F10 121,F11 122,F12 123,HOME 36,END 35,PAGEDOWN 34,PAGEUP 33,ESC 27,ENTER 13,SPACE 32"
    ��V�� = "DOWN 40,UP 38,LEFT 37,RIGHT 39"
    �r���� = "A 65,B 66,C 67,D 68,E 69,F 70,G 71,H 72,I 73,J 74,K 75,L 76,M 77,N 78,O 79,P 80,Q 81,R 82,S 83,T 84,U 85,V 86,W 87,X 88,Y 89,Z 90"
    �Ʀr�� = "0 48,1 49,2 50,3 51,4 52,5 53,6 54,7 55,8 56,9 57"
    �Ÿ��� = "~ 192,` 192,- 189,= 187,[ 219,] 221,\ 220,/ 191,? 191,< 188,> 190"
    �զX�� = ������ &","& �\���� &","& ��V�� &","& �r���� &","& �Ʀr�� &","& �Ÿ���
    //�w���˴�
    Dim �ন�j�g,�h���Ů�
    �ন�j�g = UCase(��X��)
    �h���Ů� = Replace(�ন�j�g," ","")
    //�Ѽ�
    Dim ���γr��,�r���ƶq,������X,��L,��X�ƶq
    ���γr�� = Split(�h���Ů�,",")
    �r���ƶq = UBound(���γr��) 
    //��w
    ������X = Split(�զX��,",")
    ��X�ƶq = UBound(������X)
    Dim i,k,n
    For i=0 To �r���ƶq 
        //�p����X
        For k=0 To ��X�ƶq
            ��L = Split(������X(k)," ") 
            If ��L(0) = ���γr��(i) Then 
                If �����覡 = 0 Then 
                    KeyPress Clng(��L(1)), 1
                ElseIf �����覡 = 1 Then
                    KeyPressH Clng(��L(1)), 1 
                ElseIf �����覡 = 2 Then
                    KeyPressS Clng(��L(1)), 1 
                End If 
                n = Plugin.Sys.GetTime() + �@����
                Do   
                    Delay 5
                loop Until Plugin.Sys.GetTime() >= n
                Exit For
            End If 
        Next
    Next 
End Sub 




Sub KeyList(��X��,�����覡,�@����)
    //�Ҥl�GCall lib.��L.KeyList("aA@2?"">.',/|\=-+_)(*&^QAsD",0,50)
    //�ݭn�`�N���O�G���J�@�Ӥ޸��ɡ]"�^������J�@��]""�^
    //�����覡�G�i0���q�����A1�w������A2�W�ż����j
    Dim ��X(46)
    ��X(0) ="a��A��65"
    ��X(1) ="b��B��66"
    ��X(2) ="c��C��67"
    ��X(3) ="d��D��68"
    ��X(4) ="e��E��69"
    ��X(5) ="f��F��70"
    ��X(6) ="g��G��71"
    ��X(7) ="h��H��72"
    ��X(8) ="i��I��73"
    ��X(9) ="j��J��74"
    ��X(10)="k��K��75"
    ��X(11)="l��L��76"
    ��X(12)="m��M��77"
    ��X(13)="n��N��78"
    ��X(14)="o��O��79"
    ��X(15)="p��P��80"
    ��X(16)="q��Q��81"
    ��X(17)="r��R��82"
    ��X(18)="s��S��83"
    ��X(19)="t��T��84"
    ��X(20)="u��U��85"
    ��X(21)="v��V��86"
    ��X(22)="w��W��87"
    ��X(23)="x��X��88"
    ��X(24)="y��Y��89"
    ��X(25)="z��Z��90"
    ��X(26)="`��~��192"
    ��X(27)="1��!��49"
    ��X(28)="2��@��50"
    ��X(29)="3��#��51"
    ��X(30)="4��$��52"
    ��X(31)="5��%��53"
    ��X(32)="6��^��54"
    ��X(33)="7��&��55"
    ��X(34)="8��*��56"
    ��X(35)="9��(��57"
    ��X(36)="0��)��48"
    ��X(37)="-��_��189"
    ��X(38)="=��+��187"
    ��X(39)="[��{��219"
    ��X(40)="]��}��221"
    ��X(41)="\��|��220"
    ��X(42)=";��:��186"
    ��X(43)="'��""��222"
    ��X(44)=",��<��188"
    ��X(45)=".��>��190"
    ��X(46)="/��?��191"
    //Dim KeyS()
    Dim �ƶq,�P�_,i,m,n
    �ƶq=Len(��X��)
    ReDim KeyS(�ƶq)
    For i=0 to �ƶq-1
        KeyS(i)=Mid(��X��,i+1,1)
        �P�_=False
        For n=0 to 46
            MyKeyS=Split(��X(n),"��")
            If KeyS(i)=MyKeyS(0) Then
                �P�_=True
                If �����覡 = 0 Then 
                    KeyPress Clng(MyKeyS(2)), 1
                ElseIf �����覡 = 1 Then
                    KeyPressH Clng(MyKeyS(2)), 1
                ElseIf �����覡 = 2 Then
                    KeyPressS Clng(MyKeyS(2)), 1
                End If
                Exit For
            ElseIf KeyS(i)=MyKeyS(1) Then ://�ݭn����Shift��Ӽ���
                �P�_=True
                If �����覡 = 0 Then 
                    KeyDown 16, 1
                    KeyPress Clng(MyKeyS(2)), 1
                    KeyUp 16, 1
                ElseIf �����覡 = 1 Then
                    KeyDownH 16, 1
                    KeyPressH Clng(MyKeyS(2)), 1
                    KeyUpH 16, 1
                ElseIf �����覡 = 2 Then
                    KeyDownS 16, 1
                    KeyPressS Clng(MyKeyS(2)), 1
                    KeyUpS 16, 1
                End If
                Exit For
            End If
        Next
        m = Plugin.Sys.GetTime() + �@����
        Do   
            Delay 5
        loop Until Plugin.Sys.GetTime() >= m
        If �P�_=False Then Exit Sub
    Next
End Sub

//�s�@�G�@����
//����G2009.12.24
//�ק�G2011.04.06



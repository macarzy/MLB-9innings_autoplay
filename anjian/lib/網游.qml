[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//



[General]
MacroID=eb6e2723-6175-4804-a00c-d12ea20def68




[Script]
Sub �e�x�۸�(�Ʋդ��e,�C�y���j��,�`���ɶ���)
    //��r="�J�M�A�ۤ߸۷N���o�ݤF|���ڴN�j�o�O�d���i�D�A|���F����@�ɳQ�}�a|���F�u�@�@�ɪ��M��|�e���R�P�u�ꪺ���c|�i�R�S�g�H���Ϭ�����|�Z��|�p����|�ڭ̬O����b�Ȫe�������b��.�լ}..�p"
    //�Ʋդ��e=Split(��r,"|")
    //�Ҥl�GCall lib.����.�e�x�۸�(�Ʋդ��e,1,5)
    Dim �ƶq,i
    Rem �}�l
    For �ƶq=0 To UBound(�Ʋդ��e)
        SayString �Ʋդ��e(�ƶq)
        KeyPressS 13,1
        For i=1 To �C�y���j��
            Delay 1000
        Next
    Next 
    For i=1 To �`���ɶ���
        Delay 1000
    Next
    Goto �}�l
End Sub
Sub ��κ��}�u(�����Ix����,�����Iy����,���I���Z��,���)
    //�Ҥl�GCall lib.����.��κ��}�u(400,300,20,20)
    Dim x,y,v,i,j,k
    x=�����Ix����: y=�����Iy����
    //�]�m2�I���Z��
    v=���I���Z��
    i=1
    For ���
        j=0:k=v
        For 2
            For i
                x=x+j:y=y+k
                MoveTo x,y
                Delay 10
                //LeftClick 1
            Next 
            j=v:k=0
        Next 
        i=i+1:v=v*(-1)
    Next
End Sub
Sub ��κ��}�u(�����Ix����,�����Iy����,���I���Z��,���W�b�|,���)
    //�Ҥl�GCall lib.����.��κ��}�u(400,300,20,20,20)
    Dim x0,y0,rr,l,n,r,x,y
    //�]�m��ߧ���
    x0=�����Ix����:y0=�����Iy����
    //�]�m���W�b�|
    rr=���W�b�|
    //�]�m�I���Z
    l=���I���Z��
    //��l�ƨ���
    n=0
    //�]�m�Ĥ@��b�|
    r=30
    //�]�m�e����
    For ���
        While n<3.1415926*2
            //�e�ꤽ��
            x=x0+r*cos(n)
            y=y0-r*sin(n)
            MoveTo x,y 
            Delay 10
            //LeftClick 1
            //l/r�G�I�Z���H�b�|�A���o2�I���۹��ߪ�����
            //�঳�ı�����I���K�סA2�I�����Z���N�������
            n=n+l/r
        Wend 
        //�e���@��᭫�m����
        n=0
        //�e���@���b�|���Wrr
        r=r+rr
    Next 
End Sub
Sub ��꺥�}�u(�����Ix����,�����Iy����,���I���Z��,��нd��)
    //�Ҥl�GCall lib.����.��꺥�}�u(400,300,20,300)
    //�n���ܶq
    Dim x0,y0,n,x,y,color0,cor,l,r
    //�аO���}�u��Ƕ}�l
    Rem start
    //�ܶq��l��
    x0=�����Ix����:n=1:y0=�����Iy����:x=�����Ix����
    //�]�m2�I���Z��
    l=���I���Z��
    //���w��нd�򤣶W�L800
    While x<�����Ix����+��нd��
        //�ھں��}�u�����p��U�@���I����Шý�ȵ�x
        x=x0+4*(cos(n)+n*sin(n))
        //�ھں��}�u�����p��U�@���I���a���Шý�ȵ�y
        y=y0+3*(sin(n)-n*cos(n))
        //�b���в��ʤ��e�A����ؼ��I��Ȩý�ȵ�color0
        color0=GetPixelColor(x,y)
        //���W���ʹ��Ш�ؼ��I
        Call SetCursorPos(x,y)
        //�`������
        Delay 10
        //������в��ʫ���I��Ȩý�ȵ�color
        cor=GetPixelColor(x,y)
        //�P�@�I���o���⦸��Ȥ���A���P�A�h����H�U�����Ǹ}��
        If cor<>color0 Then
            //�������
            //LeftClick 1
            //���ǩ���
            //Delay 3000
        End If 
        //�p���e�I(x,y)����I(x0,y0)���Z��
        r=Sqr((x-x0)^2+(y-y0)^2)
        //���}�u�ѼƼW�q�A�䤤l/r�G�I�Z���H�b�|�A���o2�I���۹��ߪ�����
        //�঳�ı�����I���K�סA2�I�����Z���N�������
        n=n+l/r
    Wend 
    Goto start
    //�����}�l�A���ƺ��}�u�j��
End Sub


//�s�@�G�@����
//����G2009.12.22




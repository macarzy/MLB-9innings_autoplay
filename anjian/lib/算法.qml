[Comment]
�R�O�w�O������F8.0�����X�����s�\��
�z�i�H��ۤv�`�Ϊ���ƩM�l�{�Ǽg�b�R�O�w�����ܦh�Ӹ}���h�ե�
�R�O�w�̤j���u�լO���h�Ӹ}���@�ɤ@�өR�O�A�ק�@�B�N����ק�h�B
�ثe�R�O�w�\���٦b���շ��A�������ĳ�i�H�b������F�׾´��X�A���}�Ghttp://bbs.ajjl.cn

******�`�N�I�o�O�x�责�Ѫ��R�O�w�A�Фŭק�I�קK�H�������F�ɯŮ��л\�z���ק�C******//
******          �p�ݷs�W�R�O�w�A�i�b�R�O�w�I���k���ܡu�s�ءv�R�O�w            ******//


[General]
MacroID=3f1e0d02-3591-4461-aead-f3fdd3e337b1

[Script]
Function ���c�ƲձƧ�(�Ʋ�,��^����)
    //�Ʀr="100=A|50=B|1=C|0=D|10=E|20=F|12=G|21=H"
    //�Ʋ�=Split(�Ʀr,"|")
    //�Ҥl�GMsgbox lib.��k.���c�ƲձƧ�(�Ʋ�,0)
    //�y�k�榡�G�Ʋ�() = "�Ʀr�j�p=��"
    //�ƲաGNB_PaiXu(0) = "200=A"
    //��^�����G[0����1�̤p2�̤j]
    Dim Int_A,Int_B,Int_Num,Int_Len,A_Str,Int_Temp
    //�q�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X��l���
    Int_A = 0: Int_Num = UBound(�Ʋ�)
    For Int_A=0 TO UBound(�Ʋ�) - 1
        Int_B = Int_A
        For Int_Len=0 To Int_Num - Int_A 
            //�q�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�ھڡu=�v�������A�i��j�p�P�_
            A_Str = Split(�Ʋ�(Int_A), "="): B_Str = Split(�Ʋ�(Int_B), "=")
            If CSng(A_Str(0)) > CSng(B_Str(0)) Then
                Int_Temp = �Ʋ�(Int_A): �Ʋ�(Int_A) = �Ʋ�(Int_B): �Ʋ�(Int_B) = Int_Temp
            End If 
            Int_B = Int_B + 1
        Next 
    Next 
    If ��^���� = 0 Then
        //�q�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�Ƨǵ��G
        ���c�ƲձƧ� = Join(�Ʋ�)
    ElseIf ��^���� = 1 Then
        //�q�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X��^�̤p��
        ���c�ƲձƧ� = �Ʋ�(0)
    ElseIf ��^���� = 2 Then
        //�q�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X�X��^�̤j��
        ���c�ƲձƧ� = �Ʋ�(Int_Num)
    End If 
End Function   
Function �����r��Ť��Ҧ��Ʀr(�r�Ŧ�)
    //MsgBox lib.��k.�����r��Ť��Ҧ��Ʀr("dfghhj12dsfg3.hgk54dsfg67-45678")
    Dim TQstring    
    TQstring = "Dim rExp, shuzi" & vbCrLf    
    TQstring = TQstring & "shuzi = """"" & vbCrLf 
    TQstring = TQstring & "Set rExp = CreateObject(""VBScript.RegExp"")" & vbCrLf 
    TQstring = TQstring & "rExp.Pattern = ""\d+""" & vbCrLf    //�]�m�L�o�Ҧ�
    TQstring = TQstring & "rExp.Global = True" & vbCrLf        //�]�m�����i�Ω�
    TQstring = TQstring & "Dim Num, Nums" & vbCrLf             //�w�q�ܶq
    TQstring = TQstring & "Set Nums = rExp.Execute(""" & �r�Ŧ� & """)" & vbCrLf   //����j��  
    TQstring = TQstring & "For Each Num In Nums" & vbCrLf      //�M���ǰt���X
    TQstring = TQstring & "    shuzi = shuzi & Num" & vbCrLf  
    TQstring = TQstring & "next" & vbCrLf  
    TQstring = TQstring & "Set rExp = Nothing"
    Execute TQstring 
    �����r��Ť��Ҧ��Ʀr = shuzi
End Function
Function �����~�r�������r��(����~�r)
    //�Ҥl�GMsgBox lib.��k.�����~�r�������r��("�ڬO�@�����I")
    Dim �������r��,��Ӻ~�r,�~�r�s�X,���r��,i
    �������r��=""
    For i=0 To Len(����~�r)-1
        ��Ӻ~�r=Mid(����~�r,i+1,1)    
        �~�r�s�X = 65536 + Asc(��Ӻ~�r)
        ���r�� = ""
        If (�~�r�s�X >= 45217 And �~�r�s�X <= 45252)
            ���r�� = "A"
        ElseIf (�~�r�s�X >= 45253 And �~�r�s�X <= 45760)
            ���r�� = "B"
        ElseIf (�~�r�s�X >= 45761 And �~�r�s�X <= 46317)
            ���r�� = "C"
        ElseIf (�~�r�s�X >= 46318 And �~�r�s�X <= 46825)
            ���r�� = "D"
        ElseIf (�~�r�s�X >= 46826 And �~�r�s�X <= 47009)
            ���r�� = "E"
        ElseIf (�~�r�s�X >= 47010 And �~�r�s�X <= 47296)
            ���r�� = "F"
        ElseIf (�~�r�s�X >= 47297 And �~�r�s�X <= 47613)   
            ���r�� = "G"  
        ElseIf (�~�r�s�X >= 47614 And �~�r�s�X <= 48118)   
            ���r�� = "H"  
        ElseIf (�~�r�s�X >= 48119 And �~�r�s�X <= 49061)   
            ���r�� = "J"  
        ElseIf (�~�r�s�X >= 49062 And �~�r�s�X <= 49323)   
            ���r�� = "K"  
        ElseIf (�~�r�s�X >= 49324 And �~�r�s�X <= 49895)   
            ���r�� = "L"  
        ElseIf (�~�r�s�X >= 49896 And �~�r�s�X <= 50370)   
            ���r�� = "M"  
        ElseIf (�~�r�s�X >= 50371 And �~�r�s�X <= 50613)   
            ���r�� = "N"  
        ElseIf (�~�r�s�X >= 50614 And �~�r�s�X <= 50621)   
            ���r�� = "O"  
        ElseIf (�~�r�s�X >= 50622 And �~�r�s�X <= 50905)   
            ���r�� = "P"  
        ElseIf (�~�r�s�X >= 50906 And �~�r�s�X <= 51386)   
            ���r�� = "Q"  
        ElseIf (�~�r�s�X >= 51387 And �~�r�s�X <= 51445)   
            ���r�� = "R"  
        ElseIf (�~�r�s�X >= 51446 And �~�r�s�X <= 52217)   
            ���r�� = "S"  
        ElseIf (�~�r�s�X >= 52218 And �~�r�s�X <= 52697)   
            ���r�� = "T"  
        ElseIf (�~�r�s�X >= 52698 And �~�r�s�X <= 52979)   
            ���r�� = "W"  
        ElseIf (�~�r�s�X >= 52980 And �~�r�s�X <= 53640)   
            ���r�� = "X"  
        ElseIf (�~�r�s�X >= 53689 And �~�r�s�X <= 54480)   
            ���r�� = "Y"  
        ElseIf (�~�r�s�X >= 54481 And �~�r�s�X <= 55289)   
            ���r�� = "Z"  
        EndIf  
        �������r�� = �������r�� & ���r��
    Next
    �����~�r�������r�� = �������r��
End Function     
Function ����~�r�����(����~�r)
    //�Ҥl�GMsgBox lib.��k.����~�r�����("�ڬO�@�����I")
    Dim �����w�X(395)
    �����w�X(0) = "A=-20319"
    �����w�X(1) = "Ai=-20317"
    �����w�X(2) = "An=-20304"
    �����w�X(3) = "Ang=-20295"
    �����w�X(4) = "Ao=-20292"

    �����w�X(5) = "Ba=-20283"
    �����w�X(6) = "Bai=-20265"
    �����w�X(7) = "Ban=-20257"
    �����w�X(8) = "Bang=-20242"
    �����w�X(9) = "Bao=-20230"
    �����w�X(10) = "Bei=-20051"
    �����w�X(11) = "Ben=-20036"
    �����w�X(12) = "Beng=-20032"
    �����w�X(13) = "Bi=-20026"
    �����w�X(14) = "Bian=-20002"
    �����w�X(15) = "Biao=-19990"
    �����w�X(16) = "Bie=-19986"
    �����w�X(17) = "Bin=-19982"
    �����w�X(18) = "Bing=-19976"
    �����w�X(19) = "Bo=-19805"
    �����w�X(20) = "Bu=-19784"

    �����w�X(21) = "Ca=-19775"
    �����w�X(22) = "Cai=-19774"
    �����w�X(23) = "Can=-19763"
    �����w�X(24) = "Cang=-19756"
    �����w�X(25) = "Cao=-19751"
    �����w�X(26) = "Ce=-19746"
    �����w�X(27) = "Ceng=-19741"
    �����w�X(28) = "Cha=-19739"
    �����w�X(29) = "Chai=-19728"
    �����w�X(30) = "Chan=-19725"
    �����w�X(31) = "Chang=-19715"
    �����w�X(32) = "Chao=-19540"
    �����w�X(33) = "Che=-19531"
    �����w�X(34) = "Chen=-19525"
    �����w�X(35) = "Cheng=-19515"
    �����w�X(36) = "Chi=-19500"
    �����w�X(37) = "Chong=-19484"
    �����w�X(38) = "Chou=-19479"
    �����w�X(39) = "Chu=-19467"
    �����w�X(40) = "Chuai=-19289"
    �����w�X(41) = "Chuan=-19288"
    �����w�X(42) = "Chuang=-19281"
    �����w�X(43) = "Chui=-19275"
    �����w�X(44) = "Chun=-19270"
    �����w�X(45) = "Chuo=-19263"
    �����w�X(46) = "Ci=-19261"
    �����w�X(47) = "Cong=-19249"
    �����w�X(48) = "Cou=-19243"
    �����w�X(49) = "Cu=-19242"
    �����w�X(50) = "Cuan=-19238"
    �����w�X(51) = "Cui=-19235"
    �����w�X(52) = "Cun=-19227"
    �����w�X(53) = "Cuo=-19224"

    �����w�X(54) = "Da=-19218"
    �����w�X(55) = "Dai=-19212"
    �����w�X(56) = "Dan=-19038"
    �����w�X(57) = "Dang=-19023"
    �����w�X(58) = "Dao=-19018"
    �����w�X(59) = "De=-19006"
    �����w�X(60) = "Deng=-19003"
    �����w�X(61) = "Di=-18996"
    �����w�X(62) = "Dian=-18977"
    �����w�X(63) = "Diao=-18961"
    �����w�X(64) = "Die=-18952"
    �����w�X(65) = "Ding=-18783"
    �����w�X(66) = "Diu=-18774"
    �����w�X(67) = "Dong=-18773"
    �����w�X(68) = "Dou=-18763"
    �����w�X(69) = "Du=-18756"
    �����w�X(70) = "Duan=-18741"
    �����w�X(71) = "Dui=-18735"
    �����w�X(72) = "Dun=-18731"
    �����w�X(73) = "Duo=-18722"

    �����w�X(74) = "E=-18710"
    �����w�X(75) = "En=-18697"
    �����w�X(76) = "Er=-18696"

    �����w�X(77) = "Fa=-18526"
    �����w�X(78) = "Fan=-18518"
    �����w�X(79) = "Fang=-18501"
    �����w�X(80) = "Fei=-18490"
    �����w�X(81) = "Fen=-18478"
    �����w�X(82) = "Feng=-18463"
    �����w�X(83) = "Fo=-18448"
    �����w�X(84) = "Fou=-18447"
    �����w�X(85) = "Fu=-18446"

    �����w�X(86) = "Ga=-18239"
    �����w�X(87) = "Gai=-18237"
    �����w�X(88) = "Gan=-18231"
    �����w�X(89) = "Gang=-18220"
    �����w�X(90) = "Gao=-18211"
    �����w�X(91) = "Ge=-18201"
    �����w�X(92) = "Gei=-18184"
    �����w�X(93) = "Gen=-18183"
    �����w�X(94) = "Geng=-18181"
    �����w�X(95) = "Gong=-18012"
    �����w�X(96) = "Gou=-17997"
    �����w�X(97) = "Gu=-17988"
    �����w�X(98) = "Gua=-17970"
    �����w�X(99) = "Guai=-17964"
    �����w�X(100) = "Guan=-17961"
    �����w�X(101) = "Guang=-17950"
    �����w�X(102) = "Gui=-17947"
    �����w�X(103) = "Gun=-17931"
    �����w�X(104) = "Guo=-17928"

    �����w�X(105) = "Ha=-17922"
    �����w�X(106) = "Hai=-17759"
    �����w�X(107) = "Han=-17752"
    �����w�X(108) = "Hang=-17733"
    �����w�X(109) = "Hao=-17730"
    �����w�X(110) = "He=-17721"
    �����w�X(111) = "Hei=-17703"
    �����w�X(112) = "Hen=-17701"
    �����w�X(113) = "Heng=-17697"
    �����w�X(114) = "Hong=-17692"
    �����w�X(115) = "Hou=-17683"
    �����w�X(116) = "Hu=-17676"
    �����w�X(117) = "Hua=-17496"
    �����w�X(118) = "Huai=-17487"
    �����w�X(119) = "Huan=-17482"
    �����w�X(120) = "Huang=-17468"
    �����w�X(121) = "Hui=-17454"
    �����w�X(122) = "Hun=-17433"
    �����w�X(123) = "Huo=-17427"

    �����w�X(124) = "Ji=-17417"
    �����w�X(125) = "Jia=-17202"
    �����w�X(126) = "Jian=-17185"
    �����w�X(127) = "Jiang=-16983"
    �����w�X(128) = "Jiao=-16970"
    �����w�X(129) = "Jie=-16942"
    �����w�X(130) = "Jin=-16915"
    �����w�X(131) = "Jing=-16733"
    �����w�X(132) = "Jiong=-16708"
    �����w�X(133) = "Jiu=-16706"
    �����w�X(134) = "Ju=-16689"
    �����w�X(135) = "Juan=-16664"
    �����w�X(136) = "Jue=-16657"
    �����w�X(137) = "Jun=-16647"

    �����w�X(138) = "Ka=-16474"
    �����w�X(139) = "Kai=-16470"
    �����w�X(140) = "Kan=-16465"
    �����w�X(141) = "Kang=-16459"
    �����w�X(142) = "Kao=-16452"
    �����w�X(143) = "Ke=-16448"
    �����w�X(144) = "Ken=-16433"
    �����w�X(145) = "Keng=-16429"
    �����w�X(146) = "Kong=-16427"
    �����w�X(147) = "Kou=-16423"
    �����w�X(148) = "Ku=-16419"
    �����w�X(149) = "Kua=-16412"
    �����w�X(150) = "Kuai=-16407"
    �����w�X(151) = "Kuan=-16403"
    �����w�X(152) = "Kuang=-16401"
    �����w�X(153) = "Kui=-16393"
    �����w�X(154) = "Kun=-16220"
    �����w�X(155) = "Kuo=-16216"

    �����w�X(156) = "La=-16212"
    �����w�X(157) = "Lai=-16205"
    �����w�X(158) = "Lan=-16202"
    �����w�X(159) = "Lang=-16187"
    �����w�X(160) = "Lao=-16180"
    �����w�X(161) = "Le=-16171"
    �����w�X(162) = "Lei=-16169"
    �����w�X(163) = "Leng=-16158"
    �����w�X(164) = "Li=-16155"
    �����w�X(165) = "Lia=-15959"
    �����w�X(166) = "Lian=-15958"
    �����w�X(167) = "Liang=-15944"
    �����w�X(168) = "Liao=-15933"
    �����w�X(169) = "Lie=-15920"
    �����w�X(170) = "Lin=-15915"
    �����w�X(171) = "Ling=-15903"
    �����w�X(172) = "Liu=-15889"
    �����w�X(173) = "Long=-15878"
    �����w�X(174) = "Lou=-15707"
    �����w�X(175) = "Lu=-15701"
    �����w�X(176) = "Lv=-15681"
    �����w�X(177) = "Luan=-15667"
    �����w�X(178) = "Lue=-15661"
    �����w�X(179) = "Lun=-15659"
    �����w�X(180) = "Luo=-15652"

    �����w�X(181) = "Ma=-15640"
    �����w�X(182) = "Mai=-15631"
    �����w�X(183) = "Man=-15625"
    �����w�X(184) = "Mang=-15454"
    �����w�X(185) = "Mao=-15448"
    �����w�X(186) = "Me=-15436"
    �����w�X(187) = "Mei=-15435"
    �����w�X(188) = "Men=-15419"
    �����w�X(189) = "Meng=-15416"
    �����w�X(190) = "Mi=-15408"
    �����w�X(191) = "Mian=-15394"
    �����w�X(192) = "Miao=-15385"
    �����w�X(193) = "Mie=-15377"
    �����w�X(194) = "Min=-15375"
    �����w�X(195) = "Ming=-15369"
    �����w�X(196) = "Miu=-15363"
    �����w�X(197) = "Mo=-15362"
    �����w�X(198) = "Mou=-15183"
    �����w�X(199) = "Mu=-15180"

    �����w�X(200) = "Na=-15165"
    �����w�X(201) = "Nai=-15158"
    �����w�X(202) = "Nan=-15153"
    �����w�X(203) = "Nang=-15150"
    �����w�X(204) = "Nao=-15149"
    �����w�X(205) = "Ne=-15144"
    �����w�X(206) = "Nei=-15143"
    �����w�X(207) = "Nen=-15141"
    �����w�X(208) = "Neng=-15140"
    �����w�X(209) = "Ni=-15139"
    �����w�X(210) = "Nian=-15128"
    �����w�X(211) = "Niang=-15121"
    �����w�X(212) = "Niao=-15119"
    �����w�X(213) = "Nie=-15117"
    �����w�X(214) = "Nin=-15110"
    �����w�X(215) = "Ning=-15109"
    �����w�X(216) = "Niu=-14941"
    �����w�X(217) = "Nong=-14937"
    �����w�X(218) = "Nu=-14933"
    �����w�X(219) = "Nv=-14930"
    �����w�X(220) = "Nuan=-14929"
    �����w�X(221) = "Nue=-14928"
    �����w�X(222) = "Nuo=-14926"

    �����w�X(223) = "O=-14922"
    �����w�X(224) = "Ou=-14921"

    �����w�X(225) = "Pa=-14914"
    �����w�X(226) = "Pai=-14908"
    �����w�X(227) = "Pan=-14902"
    �����w�X(228) = "Pang=-14894"
    �����w�X(229) = "Pao=-14889"
    �����w�X(230) = "Pei=-14882"
    �����w�X(231) = "Pen=-14873"
    �����w�X(232) = "Peng=-14871"
    �����w�X(233) = "Pi=-14857"
    �����w�X(234) = "Pian=-14678"
    �����w�X(235) = "Piao=-14674"
    �����w�X(236) = "Pie=-14670"
    �����w�X(237) = "Pin=-14668"
    �����w�X(238) = "Ping=-14663"
    �����w�X(239) = "Po=-14654"
    �����w�X(240) = "Pu=-14645"

    �����w�X(241) = "Qi=-14630"
    �����w�X(242) = "Qia=-14594"
    �����w�X(243) = "Qian=-14429"
    �����w�X(244) = "Qiang=-14407"
    �����w�X(245) = "Qiao=-14399"
    �����w�X(246) = "Qie=-14384"
    �����w�X(247) = "Qin=-14379"
    �����w�X(248) = "Qing=-14368"
    �����w�X(249) = "Qiong=-14355"
    �����w�X(250) = "Qiu=-14353"
    �����w�X(251) = "Qu=-14345"
    �����w�X(252) = "Quan=-14170"
    �����w�X(253) = "Que=-14159"
    �����w�X(254) = "Qun=-14151"

    �����w�X(255) = "Ran=-14149"
    �����w�X(256) = "Rang=-14145"
    �����w�X(257) = "Rao=-14140"
    �����w�X(258) = "Re=-14137"
    �����w�X(259) = "Ren=-14135"
    �����w�X(260) = "Reng=-14125"
    �����w�X(261) = "Ri=-14123"
    �����w�X(262) = "Rong=-14122"
    �����w�X(263) = "Rou=-14112"
    �����w�X(264) = "Ru=-14109"
    �����w�X(265) = "Ruan=-14099"
    �����w�X(266) = "Rui=-14097"
    �����w�X(267) = "Run=-14094"
    �����w�X(268) = "Ruo=-14092"

    �����w�X(269) = "Sa=-14090"
    �����w�X(270) = "Sai=-14087"
    �����w�X(271) = "San=-14083"
    �����w�X(272) = "Sang=-13917"
    �����w�X(273) = "Sao=-13914"
    �����w�X(274) = "Se=-13910"
    �����w�X(275) = "Sen=-13907"
    �����w�X(276) = "Seng=-13906"
    �����w�X(277) = "Sha=-13905"
    �����w�X(278) = "Shai=-13896"
    �����w�X(279) = "Shan=-13894"
    �����w�X(280) = "Shang=-13878"
    �����w�X(281) = "Shao=-13870"
    �����w�X(282) = "She=-13859"
    �����w�X(283) = "Shen=-13847"
    �����w�X(284) = "Sheng=-13831"
    �����w�X(285) = "Shi=-13658"
    �����w�X(286) = "Shou=-13611"
    �����w�X(287) = "Shu=-13601"
    �����w�X(288) = "Shua=-13406"
    �����w�X(289) = "Shuai=-13404"
    �����w�X(290) = "Shuan=-13400"
    �����w�X(291) = "Shuang=-13398"
    �����w�X(292) = "Shui=-13395"
    �����w�X(293) = "Shun=-13391"
    �����w�X(294) = "Shuo=-13387"
    �����w�X(295) = "Si=-13383"
    �����w�X(296) = "Song=-13367"
    �����w�X(297) = "Sou=-13359"
    �����w�X(298) = "Su=-13356"
    �����w�X(299) = "Suan=-13343"
    �����w�X(300) = "Sui=-13340"
    �����w�X(301) = "Sun=-13329"
    �����w�X(302) = "Suo=-13326"

    �����w�X(303) = "Ta=-13318"
    �����w�X(304) = "Tai=-13147"
    �����w�X(305) = "Tan=-13138"
    �����w�X(306) = "Tang=-13120"
    �����w�X(307) = "Tao=-13107"
    �����w�X(308) = "Te=-13096"
    �����w�X(309) = "Teng=-13095"
    �����w�X(310) = "Ti=-13091"
    �����w�X(311) = "Tian=-13076"
    �����w�X(312) = "Tiao=-13068"
    �����w�X(313) = "Tie=-13063"
    �����w�X(314) = "Ting=-13060"
    �����w�X(315) = "Tong=-12888"
    �����w�X(316) = "Tou=-12875"
    �����w�X(317) = "Tu=-12871"
    �����w�X(318) = "Tuan=-12860"
    �����w�X(319) = "Tui=-12858"
    �����w�X(320) = "Tun=-12852"
    �����w�X(321) = "Tuo=-12849"

    �����w�X(322) = "Wa=-12838"
    �����w�X(323) = "Wai=-12831"
    �����w�X(324) = "Wan=-12829"
    �����w�X(325) = "Wang=-12812"
    �����w�X(326) = "Wei=-12802"
    �����w�X(327) = "Wen=-12607"
    �����w�X(328) = "Weng=-12597"
    �����w�X(329) = "Wo=-12594"
    �����w�X(330) = "Wu=-12585"

    �����w�X(331) = "Xi=-12556"
    �����w�X(332) = "Xia=-12359"
    �����w�X(333) = "Xian=-12346"
    �����w�X(334) = "Xiang=-12320"
    �����w�X(335) = "Xiao=-12300"
    �����w�X(336) = "Xie=-12120"
    �����w�X(337) = "Xin=-12099"
    �����w�X(338) = "Xing=-12089"
    �����w�X(339) = "Xiong=-12074"
    �����w�X(340) = "Xiu=-12067"
    �����w�X(341) = "Xu=-12058"
    �����w�X(342) = "Xuan=-12039"
    �����w�X(343) = "Xue=-11867"
    �����w�X(344) = "Xun=-11861"

    �����w�X(345) = "Ya=-11847"
    �����w�X(346) = "Yan=-11831"
    �����w�X(347) = "Yang=-11798"
    �����w�X(348) = "Yao=-11781"
    �����w�X(349) = "Ye=-11604"
    �����w�X(350) = "Yi=-11589"
    �����w�X(351) = "Yin=-11536"
    �����w�X(352) = "Ying=-11358"
    �����w�X(353) = "Yo=-11340"
    �����w�X(354) = "Yong=-11339"
    �����w�X(355) = "You=-11324"
    �����w�X(356) = "Yu=-11303"
    �����w�X(357) = "Yuan=-11097"
    �����w�X(358) = "Yue=-11077"
    �����w�X(359) = "Yun=-11067"

    �����w�X(360) = "Za=-11055"
    �����w�X(361) = "Zai=-11052"
    �����w�X(362) = "Zan=-11045"
    �����w�X(363) = "Zang=-11041"
    �����w�X(364) = "Zao=-11038"
    �����w�X(365) = "Ze=-11024"
    �����w�X(366) = "Zei=-11020"
    �����w�X(367) = "Zen=-11019"
    �����w�X(368) = "Zeng=-11018"
    �����w�X(369) = "Zha=-11014"
    �����w�X(370) = "Zhai=-10838"
    �����w�X(371) = "Zhan=-10832"
    �����w�X(372) = "Zhang=-10815"
    �����w�X(373) = "Zhao=-10800"
    �����w�X(374) = "Zhe=-10790"
    �����w�X(375) = "Zhen=-10780"
    �����w�X(376) = "Zheng=-10764"
    �����w�X(377) = "Zhi=-10587"
    �����w�X(378) = "Zhong=-10544"
    �����w�X(379) = "Zhou=-10533"
    �����w�X(380) = "Zhu=-10519"
    �����w�X(381) = "Zhua=-10331"
    �����w�X(382) = "Zhuai=-10329"
    �����w�X(383) = "Zhuan=-10328"
    �����w�X(384) = "Zhuang=-10322"
    �����w�X(385) = "Zhui=-10315"
    �����w�X(386) = "Zhun=-10309"
    �����w�X(387) = "Zhuo=-10307"
    �����w�X(388) = "Zi=-10296"
    �����w�X(389) = "Zong=-10281"
    �����w�X(390) = "Zou=-10274"
    �����w�X(391) = "Zu=-10270"
    �����w�X(392) = "Zuan=-10262"
    �����w�X(393) = "Zui=-10260"
    �����w�X(394) = "Zun=-10256"
    �����w�X(395) = "Zuo=-10254"
    �~�r���� = ""
    Dim i,�~�r�X,����
    For i = 1 To Len(����~�r) 
        �~�r�X = Asc(Mid(����~�r, i, 1))
        If �~�r�X > 0 And �~�r�X < 160 Then
            //���b�����d���٭�Ÿ�
            ���� = Chr(�~�r�X)
        Else
            If �~�r�X < -20319 Or �~�r�X > -10247 Then
                //���b�����d���٭�Ÿ�
                ���� = Chr(�~�r�X)
            Else
                Dim ����,n,�~�r����
                For n = UBound(�����w�X) To 0 Step -1
                    ���� = Split(�����w�X(n), "=")
                    If CLng(����(1)) <= �~�r�X Then Exit For
                Next
                ���� = ����(0)
            End If
        End If
        �~�r���� = �~�r���� & ����
    Next
    ����~�r����� = �~�r����
End Function
Function �H���r�Ŧ�(���)
    //�Ҥl�GMsgbox lib.��k.�H���r�Ŧ�(16)
    Dim i,��m,�r�Ŧ�,�r��
    �r��="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    �r�Ŧ�=""
    For i=0 To ��� - 1
        Randomize
        ��m = Int((Len(�r��) * Rnd) + 1)
        �r�Ŧ� = �r�Ŧ� & Mid(�r��,��m,1)
    Next
    �H���r�Ŧ� = �r�Ŧ�
End Function 
Function �H���Ʀr��(���)
    //�Ҥl�GMsgbox lib.��k.�H���Ʀr��(16)
    Dim i,��m,�Ʀr,�Ʀr��
    �Ʀr="0123456789"
    �Ʀr��=""
    For i=0 To ��� - 1
        Randomize
        ��m = Int((Len(�Ʀr) * Rnd) + 1)
        �Ʀr�� = �Ʀr�� & Mid(�Ʀr,��m,1)
    Next
    �H���Ʀr�� = �Ʀr��
End Function 
Function �H�����m�W()
    //�Ҥl�GMsgbox lib.��k.�H�����m�W()
    Dim a,b,c,nei,zhu
    Randomize
    a = CInt((102) * Rnd)
    b = CInt((85) * Rnd)
    c = CInt((59) * Rnd)
    nei = ""
    Select Case a
    Case 1
        zhu = "��"
    Case 2
        zhu = "��"
    Case 3
        zhu = "��"
    Case 4
        zhu = "��"
    Case 5
        zhu = "�J"
    Case 6
        zhu = "��"
    Case 7
        zhu = "��"
    Case 8
        zhu = "�K"
    Case 9
        zhu = "��"
    Case 10
        zhu = "�^"
    Case 11
        zhu = "��"
    Case 12
        zhu = "��"
    Case 13
        zhu = "�Q"
    Case 14
        zhu = "��"
    Case 15
        zhu = "�F"
    Case 16
        zhu = "��"
    Case 17
        zhu = "��"
    Case 18
        zhu = "�L"
    Case 19
        zhu = "��"
    Case 20
        zhu = "��"
    Case 21
        zhu = "�Q"
    Case 22
        zhu = "��"
    Case 23
        zhu = "�q"
    Case 24
        zhu = "����"
    Case 25
        zhu = "��"
    Case 26
        zhu = "�J"
    Case 27
        zhu = "�q"
    Case 28
        zhu = "�Z"
    Case 29
        zhu = "��"
    Case 30
        zhu = "��"
    Case 31
        zhu = "��"
    Case 32
        zhu = "�L�J"
    Case 33
        zhu = "�Ѹ�"
    Case 34
        zhu = "�W�x"
    Case 35
        zhu = "�ڶ�"
    Case 36
        zhu = "�L��"
    Case 37
        zhu = "���]"
    Case 38
        zhu = "��"
    Case 39
        zhu = "��"
    Case 40
        zhu = "�^"
    Case 41
        zhu = "��"
    Case 42
        zhu = "��"
    Case 43
        zhu = "�J"
    Case 44
        zhu = "�E"
    Case 45
        zhu = "�E"
    Case 46
        zhu = "��"
    Case 47
        zhu = "��"
    Case 48
        zhu = "��"
    Case 49
        zhu = "��"
    Case 50
        zhu = "�P"
    Case 51
        zhu = "�q��"
    Case 52
        zhu = "��"
    Case 53
        zhu = "��"
    Case 54
        zhu = "��"
    Case 55
        zhu = "��"
    Case 56
        zhu = "��"
    Case 57
        zhu = "��"
    Case 58
        zhu = "�p"
    Case 59
        zhu = "�S"
    Case 60
        zhu = "�p"
    Case 61
        zhu = "�Q"
    Case 62
        zhu = "�w"
    Case 63
        zhu = "��"
    Case 64
        zhu = "�R"
    Case 65
        zhu = "��"
    Case 66
        zhu = "��"
    Case 67
        zhu = "�j"
    Case 68
        zhu = "�I"
    Case 69
        zhu = "�Z"
    Case 70
        zhu = "�{"
    Case 71
        zhu = "��"
    Case 72
        zhu = "��"
    Case 73
        zhu = "�`"
    Case 74
        zhu = "��"
    Case 75
        zhu = "�C"
    Case 76
        zhu = "�u"
    Case 77
        zhu = "�f"
    Case 78
        zhu = "�O"
    Case 79
        zhu = "�S"
    Case 80
        zhu = "��"
    Case 81
        zhu = "��"
    Case 82
        zhu = "��"
    Case 83
        zhu = "�N"
    Case 84
        zhu = "��"
    Case 85
        zhu = "�_"
    Case 86
        zhu = "�d"
    Case 87
        zhu = "�]"
    Case 88
        zhu = "�s"
    Case 89
        zhu = "��"
    Case 90
        zhu = "��"
    Case 91
        zhu = "�O"
    Case 92
        zhu = "��"
    Case 93
        zhu = "��"
    Case 94
        zhu = "��"
    Case 95
        zhu = "��"
    Case 96
        zhu = "�}"
    Case 97
        zhu = "�v"
    Case 98
        zhu = "��"
    Case 99
        zhu = "�L"
    Case 100
        zhu = "�K"
    Case 101
        zhu = "��"
    Case 102
        zhu = "��"
    End Select
    nei = nei & zhu
    Select Case b
    Case 1
        zhu = "�p"
    Case 2
        zhu = "�Y"
    Case 3
        zhu = "��"
    Case 4
        zhu = "�X"
    Case 5
        zhu = "�Z"
    Case 6
        zhu = "�@"
    Case 7
        zhu = "��"
    Case 8
        zhu = "��"
    Case 9
        zhu = "��"
    Case 10
        zhu = "��"
    Case 11
        zhu = "�B"
    Case 12
        zhu = "��"
    Case 13
        zhu = "��"
    Case 14
        zhu = "�g"
    Case 15
        zhu = "��"
    Case 16
        zhu = "��"
    Case 17
        zhu = "�z"
    Case 18
        zhu = "��"
    Case 19
        zhu = "��"
    Case 20
        zhu = "�E"
    Case 21
        zhu = "��"
    Case 22
        zhu = "��"
    Case 23
        zhu = "�g"
    Case 24
        zhu = "��"
    Case 25
        zhu = "��"
    Case 26
        zhu = "�N"
    Case 27
        zhu = "�a"
    Case 28
        zhu = "��"
    Case 29
        zhu = "�Q"
    Case 30
        zhu = "��"
    Case 31
        zhu = "��"
    Case 32
        zhu = "�^"
    Case 33
        zhu = "�D"
    Case 34
        zhu = "��"
    Case 35
        zhu = "�m"
    Case 36
        zhu = "��"
    Case 37
        zhu = "�X"
    Case 38
        zhu = "��"
    Case 39
        zhu = "��"
    Case 40
        zhu = "�u"
    Case 41
        zhu = "��"
    Case 42
        zhu = "��"
    Case 43
        zhu = "��"
    Case 44
        zhu = "�@"
    Case 45
        zhu = "��"
    Case 46
        zhu = "�M"
    Case 47
        zhu = "��"
    Case 48
        zhu = "�T"
    Case 49
        zhu = "�E"
    Case 50
        zhu = "�t"
    Case 51
        zhu = "��"
    Case 52
        zhu = "��"
    Case 53
        zhu = "��"
    Case 54
        zhu = "��"
    Case 55
        zhu = "�R"
    Case 56
        zhu = "��"
    Case 57
        zhu = "��"
    Case 58
        zhu = "��"
    Case 59
        zhu = "�}"
    Case 60
        zhu = "�d"
    Case 61
        zhu = "��"
    Case 62
        zhu = "�b"
    Case 63
        zhu = "��"
    Case 64
        zhu = "�s"
    Case 65
        zhu = "��"
    Case 66
        zhu = "��"
    Case 67
        zhu = "��"
    Case 68
        zhu = "��"
    Case 69
        zhu = "��"
    Case 70
        zhu = "�W"
    Case 71
        zhu = "��"
    Case 72
        zhu = "�f"
    Case 73
        zhu = "��"
    Case 74
        zhu = "��"
    Case 75
        zhu = "�v"
    Case 76
        zhu = "��"
    Case 77
        zhu = "��"
    Case 78
        zhu = "�h"
    Case 79
        zhu = "��"
    Case 80
        zhu = "��"
    Case 81
        zhu = "�P"
    Case 82
        zhu = "�U"
    Case 83
        zhu = "޳"
    Case 84
        zhu = "�q"
    Case 85
        zhu = "�a"
    End Select
    nei = nei & zhu
    //�ĤT�Ӧr
    Select Case c
    Case 1
        zhu = "��"
    Case 2
        zhu = "��"
    Case 3
        zhu = "�h"
    Case 4
        zhu = "��"
    Case 5
        zhu = "�Q"
    Case 6
        zhu = "�d"
    Case 7
        zhu = "��"
    Case 8
        zhu = "��"
    Case 9
        zhu = "��"
    Case 10
        zhu = "��"
    Case 11
        zhu = "�q"
    Case 12
        zhu = "��"
    Case 13
        zhu = "��"
    Case 14
        zhu = "��"
    Case 15
        zhu = "��"
    Case 16
        zhu = "�^"
    Case 17
        zhu = "�H"
    Case 18
        zhu = "�J"
    Case 19
        zhu = "��"
    Case 20
        zhu = "�H"
    Case 21
        zhu = "�{"
    Case 22
        zhu = "�v"
    Case 23
        zhu = "�@"
    Case 24
        zhu = "��"
    Case 25
        zhu = "��"
    Case 26
        zhu = "��"
    Case 27
        zhu = "�x"
    Case 28
        zhu = "��"
    Case 29
        zhu = "��"
    Case 30
        zhu = "�P"
    Case 31
        zhu = "�K"
    Case 32
        zhu = "��"
    Case 33
        zhu = "��"
    Case 34
        zhu = "��"
    Case 35
        zhu = "��"
    Case 36
        zhu = "ظ"
    Case 37
        zhu = "�{"
    Case 38
        zhu = "��"
    Case 39
        zhu = "��"
    Case 40
        zhu = "�["
    Case 41
        zhu = "��"
    Case 42
        zhu = "�a"
    Case 43
        zhu = "�f"
    Case 44
        zhu = "�a"
    Case 45
        zhu = "�f"
    Case 46
        zhu = "��"
    Case 47
        zhu = "�o"
    Case 48
        zhu = "��"
    Case 49
        zhu = "�A"
    Case 50
        zhu = "�l"
    Case 51
        zhu = "��"
    Case 52
        zhu = "�U"
    Case 53
        zhu = "�f"
    Case 54
        zhu = "��"
    Case 55
        zhu = "��"
    Case 56
        zhu = "��"
    Case 57
        zhu = "�@"
    Case 58
        zhu = "��"
    Case 59
        zhu = "�f"
    End Select
    �H�����m�W = nei & zhu
End Function
Function �o��r�Ŧꤤ�r�����ƶq(�r�Ŧ�)
    //�Ҥl�GMsgBox lib.��k.�o��r�Ŧꤤ�r�����ƶq("[email=abc@#$%de]abc@#$%de[/email]")
    Dim TQstring  
    TQstring = "Dim regEx, Matches, shuliang" & vbCrLf                    //�w�q�ܶq
    TQstring = TQstring & "Set regEx = New RegExp" & vbCrLf               //�إߥ��h��F��
    TQstring = TQstring & "regEx.Pattern = ""([a-z]{1})""" & vbCrLf       //�]�m�L�o�Ҧ�
    TQstring = TQstring & "regEx.IgnoreCase = true" & vbCrLf              //�]�m�O�_�Ϥ��r�Ťj�p�g
    TQstring = TQstring & "regEx.Global = True" & vbCrLf                  //�]�m�����i�Ω�
    TQstring = TQstring & "Set Matches = regEx.Execute(""" & �r�Ŧ� & """)" & vbCrLf   //����j��   
    TQstring = TQstring & "shuliang = Matches.count"
    Execute TQstring 
    �o��r�Ŧꤤ�r�����ƶq = shuliang  
End Function  
Function �Q���i����Q�i��(�Q���i��r�Ŧ�)
    //�Ҥl�GMsgbox lib.��k.�Q���i����Q�i��("FFFFFF")
    Dim D,H,i,Ia
    D = 0
    H = UCase(�Q���i��r�Ŧ�)
    For i = 1 To Len(H)
        Ia = Asc(Mid(H, i, 1)) - 48
        If Ia > 9 Then Ia = Ia - 7
        D = D * 16 + Ia
    Next
    �Q���i����Q�i�� = D
End Function
Sub ���k��ͦ���()
    //�Ҥl�GCall lib.��k.���k��ͦ���()
    Dim c,d 
    Dim str,s
    For c = 1 To 9 
        For d = 1 To c
            s = d & "��" & c & "=" & c * d 
            s = s & Space(6-len(s))
            str = str & s & " " 
        Next 
        str = str & vbCrlf  
    Next
    MsgBox str,0,"���k��ͦ���" 
End Sub 
Function �~�P(�r�Ŧꤺ�e)
    //�����G�i�H���ä@�y���e����
    //�Ҥl�GMsgBox lib.��k.�~�P("123��456")
    Dim ���G,�ƶq,i, j, temp
    �ƶq=Len(�r�Ŧꤺ�e)
    ReDim tt(�ƶq)
    ���G = ""
    For i = 1 To �ƶq 
        tt(i) = Mid(�r�Ŧꤺ�e,i,1)
    Next
    Randomize
    For j = 1 To �ƶq
        i = Int(�ƶq * Rnd + 1)
        temp=tt(i)        
        tt(i)=tt(j)
        tt(j)=temp 
    Next
    �~�P = Join(tt,"")
End Function
Function �P�_�O�_�b�@�����u�W(�_�Ix����,�_�Iy����,���Ix����,���Iy����,�P�_�Ix����,�P�_�Iy����)
    //�Ҥl�GMsgBox CBool(lib.��k.�P�_�O�_�b�@�����u�W(0,3,2,5,4,7))
    //�p�⤽���Gy=k*x+b 
    Dim k,b,y   
    �P�_�O�_�b�@�����u�W=False
    //�P�_���ƬO�_��0
    If ���Ix���� - �_�Ix���� = 0 or ���Iy���� - �_�Iy���� = 0 then
        //�u�O�P�_�L�O �ݽu �٬O ��u
        If (�P�_�Iy���� >= �_�Iy���� and �P�_�Iy���� <= ���Iy����) and (�P�_�Ix���� >= �_�Ix���� and �P�_�Ix���� <= ���Ix����) then
            �P�_�O�_�b�@�����u�W=True 
        End if
        Exit Function 
    End If
    k = abs(���Iy����-�_�Iy����) / abs(���Ix����-�_�Ix����) 
    b = �_�Iy���� - k * �_�Ix����
    y = k * �P�_�Ix���� + b
    //�]�|�p��X�p���I�A�ҥH�[�d��P�_
    If y>�P�_�Iy����-1 and y<�P�_�Iy����+1 Then
        �P�_�O�_�b�@�����u�W=True
    End If
End Function
Function ���׭p��(�����Ix����,�����Iy����,�����Ix����,�����Iy����)
    //�Ҥl�GMsgBox lib.��k.���׭p��(0,0,10,10)
    //�W��0�X�A�k��90�X�A�U��180�X�A����270�X
    Dim x,y,a,b
    If �����Ix����=�����Ix���� Then
        If �����Iy����>�����Iy���� Then
            //��
            ���׭p�� = 0
        Else
            //��
            ���׭p�� = 180
        End If
    ElseIf �����Iy����=�����Iy���� Then
        If �����Ix����>�����Ix���� Then
            //��
            ���׭p�� = 90
        Else
            //��
            ���׭p�� = 270
        End If
    Else
        If �����Ix����>�����Ix���� and �����Iy����>�����Iy���� Then
            //��
            b = 90
        ElseIf �����Ix����>�����Ix���� and �����Iy����<�����Iy���� Then
            //��
            b = 0
        ElseIf �����Ix����<�����Ix���� and �����Iy����<�����Iy���� Then
            //��
            b = 270
        ElseIf �����Ix����<�����Ix���� and �����Iy����>�����Iy���� Then
            //��
            b = 180
        End If
        x = abs(�����Iy���� - �����Iy����)
        y = abs(�����Ix���� - �����Ix����)
        If x>0 Then
            //1���׬���57.3
            a = Atn(y / x)
            ���׭p�� = fix(a * 57.3) + b
            //���׭p�� = fix(a/(3.14159265/180))
        End If
    End If
End Function


//�s�@�G�@����
//����G2009.12.22
//�ק�G2011.11.30





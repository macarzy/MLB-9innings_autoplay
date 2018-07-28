[General]
SyntaxVersion=2
BeginHotkey=103
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=96
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=c61cfd22-8203-498d-82a6-f155d2ab1f9e
Description=我的腳本2
Enable=1
AutoRun=0
[Repeat]
Type=1
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
//////////////////////////////////////////////////////////////
//MLB 9 Innings V 3.0.6										//
//Autoplay  v 1.0.0											//
//															//
//by LCH													//
//////////////////////////////////////////////////////////////
/*
Press "Num 5" continuully for 25mins to the "off/all/def_only" screen.
check its state, then go to the autoplay mode.
aint support lose 3 games, player skill levelup
*/
t = now
while DateDiff("n", t, now)< 25  
	KeyPress "Num 5", 1
	Delay 500
EndWhile

FindPic 384,252,1397,824,"C:\Users\USER\Desktop\按鍵精靈\pic\off_only.bmp",0.9,intX,intY
If intX > 0 And intY > 0 Then 
	KeyPress "i", 1
	Delay 3000
	KeyPress "c", 1
	Delay 2000
	
		For 10
		KeyPress "3", 1
		Delay 500
		EndFor
	
	KeyPress "4", 1

End If

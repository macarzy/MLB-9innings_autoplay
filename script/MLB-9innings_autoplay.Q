[General]
SyntaxVersion=2
BeginHotkey=112
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=c61cfd22-8203-498d-82a6-f155d2ab1f9e
Description=MLB_AUTO
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
//anjian 9													//
//MLB 9 Innings V 3.0.7										//
//Autoplay  v 2.2											//
//Bulestacks bete 4											//
//by LCH													//
//////////////////////////////////////////////////////////////
/*
Check its state, then go to the autoplay mode.
Press "Num 5" continuully for 1 mins till the "off/all/def_only" screen show up.
BS windows locate at 456,194,1479,861 for 1920*1080 LCD
*/

Do //offense only mode
FindPic 384,252,1397,824,"D:\MLB-9innings_autoplay\anjian\pic\off_only.bmp",0.7,intX,intY
If intX > 0 And intY > 0 Then 
	KeyPress "i", 1
	Delay 3000
	KeyPress "c", 1
	Delay 2000
	
		For 8
		KeyPress "3", 1
		Delay 500
		EndFor
	
	KeyPress "4", 1
	
Else 
	Exit Do
	
End If
Loop

/*Do //players of the month check
FindPic 456,194,1479,861,"D:\MLB-9innings_autoplay\anjian\pic\potm_ok.bmp",0.7,intX,intY
If intX > 0 And intY > 0 Then 
	KeyPress "p", 1
	Delay 1000
	KeyPress "p", 1
	Delay 1000
Else 
	Exit Do
	
End If
Loop*/

Do //skill upgrade check
FindPic 456,194,1479,861,"D:\MLB-9innings_autoplay\anjian\pic\skill_upgrade.bmp",0.7,intX,intY
If intX > 0 And intY > 0 Then
	KeyPress "l", 1
	Delay 2000
	KeyPress "q", 1
	Delay 1000
	KeyPress "1", 1
	Delay 1000
Else
	Exit Do 
End If
Loop

Do //unknown problem
FindPic 456,194,1479,861,"D:\MLB-9innings_autoplay\anjian\pic\prbm_check.bmp",0.7,intX,intY
If intX > 0 And intY > 0 Then
	KeyPress "u", 1
Else
	Exit Do 
End If
loop

Do //playoff next page start
FindPic 456,194,1479,861,"D:\MLB-9innings_autoplay\anjian\pic\next_page.bmp",0.7,intX,intY
If intX > 0 And intY > 0 Then 
	KeyPress "L", 1
Else 
	Exit Do
End If
Loop

Do //playoff next page start
FindPic 456,194,1479,861,"D:\MLB-9innings_autoplay\anjian\pic\next_page_playoffwin.bmp",0.7,intX,intY
If intX > 0 And intY > 0 Then 
	KeyPress "o", 1
Else 
	Exit Do
End If
loop

//speedup autoplay
t = now
while DateDiff("n", t, now)< 1 
	KeyPress "Num 5", 1
	Delay 500
EndWhile

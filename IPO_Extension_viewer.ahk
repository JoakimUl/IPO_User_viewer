{
#NoEnv
#SingleInstance force
SendMode Input
#InstallKeybdHook
#Persistent
Menu, Tray, Icon, shell32.dll, 278
SetWorkingDir %A_ScriptDir%
FileEncoding, CP28591 ; = iso-8859-1 = Avaya IP Office TFTP-filer. 
If !FileExist(IPO_Files)
	FileCreateDir, IPO_Files
	FileDelete IPO_Files\sysinfoE.txt
GuiName = IP Office User viewer
Ank = or select from this list
DND = none
If !FileExist(A_Workingdir . "/IPO_Extension_viewer.ini")
{
;msgbox pop
InputBox, ipo, IP Office IP Address, enter IP address to the PBX (must be reachable through TFTP),,640, 200
FileAppend,
(
; Inifile for IPO_Extension_viewer:
[init]
ipo = %ipo%
PosX = 10
PosY = 10

; Colors in RGB (HEX values):
[style]
StyleBG = E0DEEE
styleHead = 555566
styleText = 000000
styleBox = 0033F0
),	IPO_Extension_viewer.ini, UTF-8
} else {
}
ini = %A_Workingdir%/IPO_Extension_viewer.ini
List := ListA

iniread, PosX, %ini%, init, PosX
iniread, PosY, %ini%, init, PosY

iniread, StyleBG, %ini%, style, StyleBG
iniread, styleHead, %ini%, style, styleHead
iniread, styleText, %ini%, style, styleText
iniread, styleBox, %ini%, style, styleBox

gosub Whois
gosub GetUserList
Goto GuiStart
}

GuiStart:
{
	Gui 1: Destroy
	GW = 600
	GH = 278
	LX20 := GW - 38
	LX100 := GW - 100
	Gui 1: Color, %StyleBG%
	; Gui 1: +ToolWindow
	Gui 1: font, s10 bold c%styleBox%
	Gui 1: Add, Text, x10 y8,
	Gui 1: Add, Edit, x150 w60 y7 h16 R1 BackgroundTrans vAnka, 
	Gui 1: font, s12 norm c%styleText%
	Gui 1: Add, Text, x10 y8, %nam%
	Gui 1: font, s10
	Gui 1: Add, Button, x212 y6 h26 BackgroundTrans default gSelMan , Get user
	Gui 1: Add, DDL, x+10 yp+1 vAnk gSelDDL h600 w240, %uList%
	;Gui, Add, DDL, x180 w100 y8 vOrder gOrderN, List Order||by number|by name
	Gui 1: font, s10 Bold
	Gui 1: Add, Button, x%LX20% y6 w32 gSysInfo, IP
	
	; -------------------------- Number, Name, User Rights:	
	Gui 1: font, s12 norm underline c%styleHead%
	Gui 1: Add, GroupBox, x-2 y+2 w800 h76, 
	Gui 1: Add, Text, yp+18 x10 BackgroundTrans , ExtNbr: Full/Display Name
	Gui 1: Add, Text, YP+0 x300 BackgroundTrans, User ID:
	Gui 1: Add, Text, YP+0 x430 BackgroundTrans, User Rights group:
	
	Gui 1: font, norm s14 c%styleText%
	Gui 1: Add, Text, yp+22 x10 w250 vNamn, 
	Gui 1: font, norm s12
	Gui 1: Add, Text, yp+0 x300 w300 vuID, 
	Gui 1: Add, Text, yp+0 x430 w300 vRgt, 
	
	; -------------------------- Mobile, Voicemail, AbsText:
	Gui 1: font, s12 underline c%styleHead%
	Gui 1: Add, Text, Y+22  x10, Mobile number:%A_Space%%A_Space%%A_Space%%A_Space%%A_Space%%A_Space%
	Gui 1: Add, Text, yp+0 x140, Twin On/Off:
	Gui 1: Add, Text, yp+0 x250, Avaya Vmail - New:
	Gui 1: Add, Text, yp+0 x410, Display/Absent text:

	Gui 1: font, norm s14 vTwCol c%styleText%
	Gui 1: Add, Text, x10 yp+22 w300 vTwn, 
	Gui 1: Add, Text, x140 yp+0 w200 vTwo, 		
	Gui 1: font, norm c%styleText%
	Gui 1: Add, Text, x250 yp+0 w300 vRblMsg, 
	Gui 1: Add, Text,  x410 w300 yp+0 vAbsnt, 

	; ----------------------------- DND, Forwarding & Follow Me:
	Gui 1: font, s12 underline c%styleHead%
	Gui 1: Add, GroupBox, x-2 w800 h70, 
	Gui 1: Add, Text, yp+18 x10, DND:
	Gui 1: Add, Text, YP+0 x90, Forwarding/HV:
	Gui 1: Add, Text, YP+0 x250, Busy fwd `, NoA fwd  To:
	Gui 1: Add, Text, YP+0 x500, Follow me:

	Gui 1: font, s16 norm bold cFF0000
	Gui 1: Add, Text, YP+22 x10 w80 BackgroundTrans vDND,
	Gui 1: font, s14 norm c%styleText%
	Gui 1: Add, Text, yp+0  x90 w250 BackgroundTrans vVK, 
	Gui 1: Add, Text, YP+0 x260 w200 BackgroundTrans vBsy,
	Gui 1: Add, Text, YP+0 x330 w200 vNoA,	
	Gui 1: Add, Text, YP+0 x375 w200 vBDn,
	Gui 1: font, norm bold cFF0000
	Gui 1: Add, Text, YP+0 x500 w200 vFlw,	

	; ----------------------------- Other:	
	Gui 1: font, s12 norm underline c%styleHead%
	Gui 1: Add, Text, yp+34 x10 BackgroundTrans, Extn In/out:
	Gui 1: Add, Text, yp+0 x110 BackgroundTrans, User Licences / Applications:
	Gui 1: font, s14 norm c%styleText%
	Gui 1: Add, Text, YP+22 x10 w300 vInl, 
	Gui 1: font, s12 norm 
	Gui 1: Add, Text, YP+0 x110 w700 vFea, 
	
	Gui 1: font, norm s8 c%styleHead%
	Gui 1: Add, Text, x10 y+3, / By Joakim Ulfeldt
	Gui 1: Show, W%GW% x%PosX% y%PosY%, %GuiName%
	SetTimer, Sel, 1000
	Return
	}

SysInfo:
WinGetPos, GX, GY, GW, GH, %GuiName%
IX := GX + GW - 196
Gui 2: Destroy
Gui 2: Color, %StyleBG%
;Run Notepad.exe IPO_Files\sysinfo.txt
iniread, ipo, %ini%, init, ipo 
;sleep 1000
Gui 2: -SysMenu
Gui 2: Add, Edit, x8 w90 y10 r1 vNewIP, %ipo%
Gui 2: Add, Button, x+3 y10 gWmanager, Web Manager
Gui 2: Add, Button, x8 gChangeIP, Update IP
Gui 2: Add, Button, x+4 yp w30 gCloseIP, Cancel
Gui 2: Show, x%IX% y%GY% w190 H75, Address to the IP Office
Return

Wmanager:
{
type := si[4]
		type := StrSplit(type, " ")
		if type[2] = 500
		{
		Webmgr = https://%ipo%:8443/WebMgmtEE/WebManagement.html
		} else {
		Webmgr = https://%ipo%:7070/WebManagement/WebManagement.html
		}
Run, %Webmgr%
Return
}



Whois:
{
iniread, ipo, %ini%, init, ipo 
Run, %comspec% /c tftp.exe %ipo% GET nasystem/who_is IPO_Files\sysinfoE.txt,,Hide
	sleep 500
	If !FileExist(A_Workingdir . "\IPO_Files\sysinfoE.txt")
	{
	; SetTimer, TFTPget, Off
	;Gosub NoAnswer
	FileDelete IPO_Files\AllUserList.txt
	InputBox, ipo, IP Office IP Address, enter IP address to the PBX (must be reachable through TFTP),,640, 200
	IniWrite, %ipo%, %ini%, init, ipo
	ipo = %ipo%
	Goto Whois
	Return
	}
Return
}

GetUserList:
{ ; Get user_list. create DDL 
iniread, ipo, %ini%, init, ipo 
Gosub Whois
	fileread, sysinfo, IPO_Files\sysinfoE.txt
	si := StrSplit(sysinfo, "`""","")  ; " Omits periods.
	mac := si[1] . " " . si[2] 
	typ := si[3] . " " . si[4] 
	ver := si[9] . " " . si[10]
	nam := si[11] . " " . si[12]
	nm := strsplit(nam, A_Space)
	nam := nm[3]	
; if no mac then show text reason.
If nam = 
	{
	msgbox, 0, Check connection to %ipo%, 1. Can you ping %ipo%? `n`n2. is TFTP client installed on this PC? `n`n3. is TFTP activated in the IPO/security?,0
	exitapp
	}
; Fält 5 ger UserLicenser - beräkning se TFTP realtime IPO.rtf
RunWait, %comspec% /c tftp.exe %ipo% GET nasystem/user_list8 IPO_Files\AllUserList.txt,,Hide
dList =
ListA =
ListN =
uList =
Rad =
SN =
EX =
FN =
FileRead, dList , IPO_Files\AllUserList.txt
Sort, dList, D
; msgbox % dlist
loop Parse, dList, `n, `r
{
Rad := A_LoopField
Rad := StrSplit(Rad, ",")
SN := Rad[1]
EX := Rad[2]
FN := Rad[3]
IF (SN = "NoUser" or SN = "" )
	{
	} else {
	ListN := ListN . EX . ": " . FN . "|"
	}
If A_Index > 1235
Break
}
ListA := ListN 
Sort, ListN, D:
ListA = %Ank%||%ListA%
ListN = %Ank%||%ListN%
uList := ListA
Return
}


Sel:
{
BNs =
RunWait, %comspec% /c tftp.exe %ipo% GET nasystem/user_info3/%Ankn% IPO_Files\%Ankn%.txt,,Hide
; If !FileExist("IPO_Files\" Ankn ".txt")
;	Return	
FileRead Usr, IPO_Files\%Ankn%.txt
U := StrSplit(Usr, ",")
uID := U[1] ;
Nam := U[2] ;
Ext := U[3] ;
	Namn := Ext ": " Nam
	
BNs := U[4] + U[5] ; Busy/NoAnsw: 0=none, 1=one active 2=both active. 
Bsy := U[4] = 1 ? "Yes" : ""
NoA := U[5] = 1 ? "Yes" : ""
BDn := BNs < 1 ? "(" U[8] ")" : U[8]
BDn := U[8] > 0 ? BDn : "" 

Fwd := U[6] = 1 ? "On ▶▶ " : "off " ;   Om VK på
{ 
If U[6] = 1
	{ ; Fwd is ON
	VKcolor = c007000
	Dst := U[7] < 1 ? "NO destination!" : U[7]
	If U[7] < 1
		VKcolor = cFF0000
	}
	else
	{ ; Fwd is OFF:
	VKcolor := styleText
	Dst := U[7] < 1 ? "" : "(" U[7] ")" ; Om VK inte är på : (Dst)
	}
Fwd := U[1] < 1 ? "" : Fwd ; Visa bara Fwd om ankn är vald
Fwd := U[7] < 1 ? "" : Fwd ; Visa bara Fwd om Dst-nr finns
VK := Fwd . Dst
}
DND := U[9] > 0 ? "⛔DND" : "" ; ternary operator: If DND > 0 then store sign in DND, else store nothing in DND
Flw := U[10]> 0 ? "To: " Flw : ""

Rbl := U[12]> 0 ? " On -" : "Off -" ;
Msg := U[14]> 0 ? U[14] . " new" : "none"
Rbl := U[1] < 1 ? "" : Rbl ; Visa inget på VMail om ingen ankn visas.
Msg := U[1] < 1 ? "" : Msg 
	RblMsg := Rbl " " Msg

Atp := U[38] ;
Atx := U[39] ;
Aon := U[40] ;
Inl := U[45] ;
Rgt := U[54] ;
Twn := U[68] ;
Two := U[69] ;

Fea := ""
  If U[1] > 1
  {
  	loop Parse, dList, `n, `r
	{
	Rad := A_LoopField
	uFt := StrSplit(Rad, ",")
	uFL := uFt[1]
	If uFL = %uID%
		{
		uFea := uFt[5]
		If Mod(uFea, 2) != 0 ; if Odd number; Receptionist is activated. Remove 1 from applications sum.
			{
			Rec := "Receptionist. "
			rece += 1
			uFea := uFea - 1
			} else {
			}
			if uFea > 512 
				{
				Teams := ". Teams"
				uFea := uFea - 512
				}
			if uFea > 62
				{
				1xWc := "+Web Collab Conf" 
				uFea := uFea - 64
				}
			if uFea > 30
				{
				WpM := uFt[4]= 4 ? "+mVoIP" : ""
				uFea := uFea - 32
				}
			if uFea > 14
				{
				If uFt[4]= 4 
					WpD := "Power User(WP)"
				If uFt[4]= 5
					WpD := "Office Worker(WP)"
				If uFt[4]= 2
					WpD := "Telewrk"
				If uFt[4]= 3 
					WpD :="MobileWrk"
				uFea := uFea - 16
				}
			if uFea > 6 
				{
				RemW := uFt[4]= 1 or 2 or 3 or 5 ? ". remote" : ""
				uFea := uFea - 8
				}
			if uFea > 2 
				{
				1xTc := " +comm"
				uFea := uFea - 4
				}
			if uFea = 2
				{
				Old := ". pc"
				uFea := uFea - 2
				}
				
		OneX := uFt[6]=1 ? ". OneXp" . 1xWc : ""		
				
				
			Fea := Rec . WpD . WpM . OneX . Teams
			Rec =
			WpD =
			WpM =
			RemW =
			Old =
			Teams =
			
			}
		}
	}

{
If Atp = 1
Atp := "(type 1) On vacation until " 
If Atp = 2
Atp := "(type 2) Will be back " 
If Atp = 3
Atp := "(type 3) At Lunch until " 
If Atp = 4
Atp := "(type 4) Meeting until " 
If Atp = 5
Atp := "(type 5) Please Call " 
If Atp = 6
Atp := "(type 6) Dont disturb until " 
If Atp = 7
Atp := "(type 7) With visitors until " 
If Atp = 8
Atp := "(type 8) With customers til " 
If Atp = 9
Atp := "(type 9) Back soon " 
If Atp = 10
Atp := "(type 10) Back tomorrow " 
If Atp = 11
Atp = 
Absnt := Atp Atx
If Aon = 0
Absnt =
}
If Inl = 0
Inl = Out
If Inl = 1
Inl = In

If Twn =
	{
	Two =
	} else {
	If Two = 0
		{
		TwCol = c000000
		Two = off
		}
	If Two = 1
		{
	TwCol = cA0A080
	Two = Twin On
		}
	}
}

{	; Update the GUI
	GuiControl,,uID, %uID%
	GuiControl,,Namn, %Namn%
	GuiControl,,Rgt, %Rgt%
	GuiControl,,Twn, %Twn%
	GuiControl,,Two, %Two%
	GuiControl,,TwCol, %TwCol%
	GuiControl, +c%VKcolor% +Redraw, VK
	GuiControl,,VK, %VK%
	GuiControl,,Bsy, %Bsy%
	GuiControl,,NoA, %NoA%
	GuiControl,,BDn, %BDn%
	GuiControl,,Flw, %Flw%
	GuiControl,,RblMsg, %RblMsg%
	GuiControl,,Absnt, %Absnt%
	GuiControl,,Inl, %Inl%
	GuiControl,,DND, %DND%
	GuiControl,,Fea, %Fea%
	Fea := ""
Return
}

CloseIP:
Gui 2: Destroy
Return

ChangeIP:
Gui 2: Submit, NoHide
IniWrite, %NewIP%, %ini%, init, ipo
Gui 2: Destroy
gosub Whois
gosub GetUserList
Gui 1: Destroy
Goto GuiStart
Return

SelMan:
Gui 1: Submit, NoHide
Ankn := Anka
GuiControl,,Ank, %List%
Goto Sel

SelDDL:
Gui 1: Submit, NoHide
Ankn := StrSplit(Ank, ":")
Ankn := Ankn[1]
GuiControl,,Anka, %Ankn%
Goto Sel

OrderA:
Gui 1: Submit, Hide
Gosub GetUserList
List := ListA
Goto GuiStart
Return

OrderN:
Gui 1: Submit, Hide
Gosub GetUserList
List := ListN
Goto GuiStart
Return

GuiClose:
WinGetPos, posX, posY, , , %GuiName% 
IniWrite, %posX%, %ini%, init, posX
IniWrite, %posY%, %ini%, init, posY
	ExitApp

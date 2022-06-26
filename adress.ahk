; mit Land + letzter Änderung
; Check auf Veränderungen vor Speichern
; statt Buttons: Icons + Tooltips 
; variable Zahl Telefonnummern/Mailadressen
; zum Wählen der Tel.nr. mit Fritzbox fb_anruf.ahk aufrufen
; übergabe an Mailprogramm/Zwischenablage
; direkte Etikettendruckeransteuerung (Nur-Text)
#singleinstance force 
Menu,Tray,Icon,pifmgr.dll,16
Menu,Tray,Add,&Anzeigen,search
Menu,Tray,Default,&Anzeigen
settitlematchmode,Regex
setworkingdir,%A_ScriptDir%
if 1=
  database=%appdata%\adress.dat
groups=Firma|Privat|obsolet|alle
aPrinter:="Adressetiketten"
onexit,ende
hCurs:=DllCall("LoadCursor","UInt",NULL,"Int",32649,"UInt") 
OnMessage(0x200,"WM_MOUSEMOVE") 
regread,row1,HKCU,Software\Clients\Mail
regread,row2,HKLM,SOFTWARE\Clients\Mail\%row1%\Protocols\mailto\shell\open\command
regexmatch(row2,"\x22(.+?)\x22",row)
if strlen(row1)<7
  regexmatch(row2,"(.+?)\s",row)
Mailer:=row1

suchtext_TT=Namensbestandteil suchen
group_TT=Adressen-Gruppe wählen (Alt+g)
ButtonBearbeiten_TT=Adressdaten bearbeiten (Enter)
ButtonNeu_TT=Neuen Eintrag erstellen (Einfg)
ButtonLöschen_TT=Eintrag löschen (Entf)
Anrufen_TT=Telefonnummer anrufen (F5)
Anrufen2_TT=Telefonnummer anrufen (F5)
EMailen_TT=E-Mail schreiben (F6)
ButtonDruck_TT=Adressetikett drucken (F9)
ButtonCopy_TT=Adresse in Zwischenablage kopieren
GuiClose_TT=Programm beenden (Escape)
privat_TT=Aktiviert = privater Absender wird gedruckt`nAusgegraut = kein Absender

Gui,+Owner 
Gui,Margin,2,5
Gui,Add,Picture,Section x12 y7 w16 h-1 Icon23,c:\windows\system32\shell32.dll 
Gui,Margin,10,5
Gui,Add,Edit,ys-1 W100 H20 gsearch vsuchtext,%suchtext%
Gui,Add,Text,xs+35 ys+2,Suche 
Gui,Font,s11
Gui,Add,ListView,x2 ys+25 r30 w615 gMyListView vMyListView Sort -Multi,Name|Name 2|Straße|PLZ|Ort|Telefon|E-Mail|Fax|Internet|Kdnr|Bemerkung|Gruppe|Land|l. Änderung
f1=150|100|100|50|100|100|100|100|100|100|1|20|100|85
loop,parse,f1,|
  LV_ModifyCol(a_index,a_loopfield)
Gui,Font,s8
Gui,Add,Text,Section x160 y8,&Gruppe:
Gui,Add,ComboBox,ys-2 vgroup gsearch Choose1,%groups%
Gui,Add,Button,xp ys w1 Default Hidden gButtonBearbeiten,Bearbeiten
Gui,Add,Picture,ys-3 w20 h-1 vButtonBearbeiten gButtonBearbeiten Icon71,c:\windows\system32\shell32.dll
Gui,Add,Picture,ys-3 w20 h-1 vButtonNeu gButtonNeu Icon100,c:\windows\system32\shell32.dll
Gui,Add,Picture,ys-3 w20 h-1 vButtonLöschen gButtonLöschen Icon132,c:\windows\system32\shell32.dll
Gui,Add,Picture,ys-3 w20 h-1 vButtonCopy gButtonCopy Icon55,c:\windows\system32\shell32.dll
Gui,Add,Picture,ys-3 w20 h-1 vAnrufen gButtonAnrufen Icon197,c:\windows\system32\shell32.dll
if Mailer>
  Gui,Add,Picture,ys-3 w20 h-1 vEMailen gButtonEMail,%Mailer%
else
  Gui,Add,Picture,ys-3 w20 h-1 vEMailen gButtonEMail Icon215,c:\windows\system32\shell32.dll
Gui,Add,Picture,ys-3 w20 h-1 vButtonDruck gButtonDruck Icon137,c:\windows\system32\shell32.dll
Gui,Margin,2,2
fileread,f1,%database%
search:
Gui,Submit,NoHide
if suchtext>
  GuiControl, Hide,Static2
blockinput,on
LV_Delete()
GuiControl,Hide,MyListView 
Loop,parse,f1,`n,`r
{
  StringSplit,row,A_Loopfield,|
  stringreplace,row6,row6,¶,`n,a
  stringreplace,row7,row7,¶,`n,a
  stringreplace,row11,row11,¶,`n,a  
  if instr(row1,suchtext) and (instr(group,row12)=1 or group="alle")
    LV_Add("", row1, row2,row3,row4,row5,row6,row7,row8,row9,row10,row11,row12,row13,row14)
}
GuiControl,Show,MyListView
blockinput,off
Gui,Show,,% "Adressen (" . LV_GetCount() . ")"
LV_Modify(1,"Focus Select")
send, {Ctrl up}
send, {Shift up}
return

#ifwinactive,^Wählen$
~right::
~left::
~up::
~down::
CursorTaste=1
return

#ifwinactive,^Adressdaten$
F5::
goto,ButtonAnrufen
F6::
goto,ButtonEMail
<^>!<::return 

#ifwinactive,^Adressen
~Del::
GuiControlGet, Focus, Focus
if (Focus="SysListView321")
  goto,ButtonLöschen
return
~down::
GuiControlGet, Focus, Focus
if (Focus="Edit1")
  send,{tab}
return
F5::
goto,ButtonAnrufen
F6::
goto,ButtonEMail
F9::
goto,ButtonDruck

Ins::
ButtonNeu:
neuanlage=1

ButtonLöschen:
neuanlage:=(neuanlage=1) ? neuanlage : -1

ButtonBearbeiten:
MyListView:
if (A_GuiEvent="ColClick")
  return
neuanlage*=2 
ifwinexist,Adressdaten,Na&me 2 
{
  msgbox,,Achtung,Erst anderes Adressdaten-Fenster schließen!
  winactivate
  return
}
ControlGet, Zeile, List, Count Focused, SysListView321, Adressen
if neuanlage=2
  Loop % LV_GetCount("Col")
    Row%a_index%:=""
else
  Loop % LV_GetCount("Col") {
    LV_GetText(Row%a_index%,Zeile,a_index)
    OldRow%a_index%:=Row%a_index%
  }

if (neuanlage<>-2) {
  Gui,2:Font,s11
  EOpts:="x100 yp-3 W300",TOpts:="x12 yp+30"
  Gui,2:Add,Text,%TOpts% y10,&Name:
  Gui,2:Add,Edit,%EOpts% vname,%Row1%
  Gui,2:Add,Text,%TOpts%,Na&me 2:
  Gui,2:Add,Edit,%EOpts% vname2,%Row2%
  Gui,2:Add,Text,%TOpts%,St&raße:
  Gui,2:Add,Edit,%EOpts% vStrasse,%Row3%
  Gui,2:Add,Text,%TOpts%,&PLZ / Ort:
  Gui,2:Add,Edit,x100 yp-3 W50 vPLZ,%Row4%
  Gui,2:Add,Edit,x155 yp W245 vOrt,%Row5%
  Gui,2:Add,Text,%TOpts%,&Telefon:
  Gui,2:Add,Edit,%EOpts% R3 vTelefon,%row6%
  Gui,2:Add,Text,x12 yp+62,&E-Mail:
  Gui,2:Add,Edit,%EOpts% R2 vemail,%Row7%
  Gui,2:Add,Text,x12 yp+46,Fa&x:
  Gui,2:Add,Edit,%EOpts% vFax,%Row8%
  Gui,2:Add,Text,%TOpts%,&Internet:
  Gui,2:Add,Edit,%EOpts% vInternet,%Row9%
  Gui,2:Add,Text,%TOpts%,&Kdnr:
  Gui,2:Add,Edit,%EOpts% vKdnr,%Row10%
  Gui,2:Add,Text,%TOpts%,Bemerk&ung:
  Gui,2:Add,Edit,%EOpts% R4 vBemerkung,%Row11%
  Gui,2:Add,Text,x12 yp+78,&Land:
  Gui,2:Add,Edit,x100 yp-3 W120 vLand,%Row13%
  Gui,2:Add,Text,x230 yp+3,&Gruppe:
  Gui,2:Add,ComboBox,x285 yp-3 W115 vGruppe,%row12%||%groups%
  Gui,2:Font,s8
  Gui,2:Add,Button,x410 y8 W80 vnx1,Etiketten&druck
  Gui,2:Add,Button,x410 y36 W80 vbüwa,Etikett Bü&Wa
  Gui,2:Add,Checkbox,x417 y63 vprivat Check3,Pri&vatabs.

  Gui,2:Add,Picture,x436 y90 w20 h-1 vButtonCopy gButtonCopy Icon55,c:\windows\system32\shell32.dll
  Gui,2:Add,Picture,x430 y129 w32 h-1 vAnrufen2 gButtonAnrufen Icon197,c:\windows\system32\shell32.dll
  if Mailer>
    Gui,2:Add,Picture,x430 y179 w32 h-1 vemailen gButtonEMail,%Mailer%
  else
    Gui,2:Add,Picture,x430 y179 w32 h-1 vemailen gButtonEMail Icon215,c:\windows\system32\shell32.dll
 
  row14n:=substr(row14,5,2) . "." . substr(row14,3,2) . "." . substr(row14,1,2) . " " . substr(row14,7,2) . ":" . substr(row14,9,2)
  Gui,2:Add,Text,x410 y300 0x2,letzte Änderung:`n%row14n%
  Gui,2:Add,Button,Default x410 y346 W80 vSpeichern,&Speichern
  Gui,2:Add,Button,x410 yp+28 W80 vnx2,A&bbrechen
  Gui,2:Show,,Adressdaten
  return
}

2ButtonSpeichern:
Gui,2:Submit,NoHide
GuiControlGet, FocusedControl, FocusV
alt:="",f2:=""
stringreplace,row11,row11,`n,¶,a  
stringreplace,Bemerkung,Bemerkung,`n,¶,a
stringreplace,row6,row6,`n,¶,a
stringreplace,Telefon,Telefon,`n,¶,a
stringreplace,row7,row7,`n,¶,a
stringreplace,email,email,`n,¶,a
loop,12
  alt.= row%a_index% . "|"
alt.=row13
neu=%name%|%name2%|%Strasse%|%PLZ%|%Ort%|%Telefon%|%email%|%Fax%|%Internet%|%Kdnr%|%Bemerkung%|%Gruppe%|%Land%
if (neuanlage=-2) {
  msgbox,257,Adressdaten,Dieser Datensatz wird unwiderruflich gelöscht:`n`n%alt%
  ifmsgbox,OK
  {
    LV_Delete(zeile)
    Gui,Show,,% "Adressen (" . LV_GetCount() . ")"
    Loop,parse,f1,`n,`r       
      f2.=(A_Loopfield=alt . "|" . row14) ? "" : A_Loopfield . "`n"
    stringtrimright,f1,f2,1   
    filewrite(database,f1)
  }
} else {
  if (alt<>neu and FocusedControl<>"Speichern") 
    msgbox,35,Adressdaten,Geänderte Daten speichern?
  ifmsgbox,Cancel
    return
  Gui,2:Destroy
  ifmsgbox,No
    return
  if (neuanlage=2 and neu="||")
    return
  formattime,row14n,,yyMMddHHmm
  if (alt<>neu) {
    if neuanlage=2
      f1.=(f1="" ? "" : "`n") . neu . "|" . row14n
    else {
      Loop,parse,f1,`n,`r      
        f2.=(A_Loopfield=(alt . "|" . row14)) ? neu . "|" . row14n . "`n" : A_Loopfield . "`n"
      stringtrimright,f1,f2,1  
    }
    row14:=row14n
    filewrite(database,f1)
    settimer,list_akt,-1
  }
}
return

ButtonBeenden:
GuiEscape:
GuiClose:
ExitApp

list_akt:
stringreplace,Bemerkung,Bemerkung,¶,`n,a
stringreplace,Telefon,Telefon,¶,`n,a
stringreplace,email,email,¶,`n,a
if instr(name,suchtext) and (instr(group,gruppe)=1 or group="alle")
  if neuanlage=2
    LV_Add("Focus Select Vis",name,name2,Strasse,PLZ,Ort,Telefon,email,Fax,Internet,Kdnr,Bemerkung,Gruppe,Land,row14)
  else
    LV_Modify(Zeile,"Focus Select",name,name2,Strasse,PLZ,Ort,Telefon,email,Fax,Internet,Kdnr,Bemerkung,Gruppe,Land,row14)
else
  if neuanlage<>2 
    LV_Delete(zeile)
neuanlage=
Gui,Show,,% "Adressen (" . LV_GetCount() . ")"
return

2ButtonAbbrechen:
2GuiEscape:
2GuiClose:
Gui,2:Destroy
return

ButtonCopy:
ControlGet, Zeile, List, Count Focused, SysListView321, Adressen
Loop % LV_GetCount("Col")
  LV_GetText(Row%a_index%,Zeile,a_index)
name:=row1,name2:=row2,strasse:=row3,plz:=row4,ort:=row5,kdnr:=row10
Gui,2:Submit,NoHide
name.=name2>"" ? "`r`n" . name2 : ""
ort.=kdnr>"" ? "`r`nKd.nr.: " . kdnr : ""
clipboard=%name%`r`n%strasse%`r`n%plz% %ort%`r`n
tooltip,=== OK ===
sleep,900
tooltip
return

ButtonAnrufen:
ControlGet, Zeile, List, Count Focused, SysListView321, Adressen
Loop % LV_GetCount("Col")
  LV_GetText(Row%a_index%,Zeile,a_index)
name:=Row1,telefon:=Row6

ifwinexist,^Wählen$
{
  winactivate
  return
}
Gui,2:Submit,NoHide
CursorTaste=0
regexmatch(telefon,"([\d\x20\-/]{4,})[^\R]",tel)
if tel1=
  return
nr_anz:=1
regexmatch(telefon,"m`a)(.+)$",tel)
Gui,3:Add,Text,,Welche Nummer soll gewählt werden?`n
Gui,3:Add,Radio,x20 gNrSelect vtel Checked,%tel1%
loop,parse,telefon,`n,`r
  if a_index>1
    {
      regexmatch(A_Loopfield,"\D*([\d\s\-/]+).*",nr)
      if strlen(nr1)>3 or (strlen(nr1)>9 and substr(nr1,1,1)="0") {
        Gui,3:Add,Radio,gNrSelect,%nr%
        nr_anz++
        tel%nr_anz%:=nr
      }
    }
if (nr_anz>1) and (a_guicontrol<>"Anrufen") { 
  Gui,3:Add,Button,Default Section x12 yp+40 W80,&OK
  Gui,3:Add,Button,xp+110 ys W80,&Abbrechen
  Gui,3:Show,,Wählen
  return
}

NrSelect:
if CursorTaste=1
  return 

3ButtonOK:
Gui,3:Submit
nr:=tel%tel%
if nr>
{
  Gui,4:Add,Text,Section x12 y20,Nummer: 
  Gui,4:Add,Edit,xs ys W120 -WantReturn H20 vnr,%nr%
  Gui,4:Add,Checkbox,gclip_jn ys+3,Nummer übermitteln
  Gui,4:Add,Button,Default Section x12 W80,Ok
  Gui,4:Add,Button,xp+90 ys W80,&Abbrechen
  Gui,4:Show,,Telefonanruf %name%
  return

  4ButtonOk:
  Gui,4:Submit
  if nr>
    run,ahk.exe %a_scriptdir%\fb_anruf.ahk "%clip%%nr%",,min

  4ButtonAbbrechen:
  4GuiEscape:
  4GuiClose:
  Gui,4:Destroy
  return

  clip_jn:
  clip:=clip="" ? "#31#," : ""
  return
}
3ButtonAbbrechen:
3GuiEscape:
Gui,3:Destroy
return

ButtonEMail:
ControlGet, Zeile, List, Count Focused, SysListView321, Adressen
LV_GetText(email,Zeile,7)

Gui,2:Submit,NoHide
row1:=""
regexmatch(email,"([\w\.\-\+_]+@[\w\.\-]+)",row)
run,%Mailer% mailto:%row1%
return

ButtonDruck:
ControlGet, Zeile, List, Count Focused, SysListView321, Adressen
Loop % LV_GetCount("Col")
  LV_GetText(Row%a_index%,Zeile,a_index)
name:=Row1,strasse:=Row3,plz:=Row4,ort:=Row5,gruppe:=Row12,land:=Row13
if gruppe=B
  name:=Row2 . " " . Row1
else
  name2:=Row2
2ButtonEtikettendruck:
2ButtonEtikettBüwa:
Gui,2:Submit,NoHide
ypos:=(name2>"") ? 200 : 160 
ypos:=(a_guicontrol="büwa") ? ypos+40 : ypos
ypos:=(land>"") ? ypos+40 : ypos
if (privat=-1)
  ypos-=50
out:="! 0 100 " . ypos . " 1`n"

if (privat=1)
  out.="TEXT 1 240 0 Absender privat...`n"
else
  if privat<>-1
    out.="TEXT 1 240 0 Absender Geschäft...`n"
ypos:=(privat=-1) ? -40 : 10 
if (a_guicontrol="büwa") {
  ypos+=40,out.="TEXT 2 240 " . ypos . " BÜWA`n"
}
ypos+=40,out.="TEXT 2 240 " . ypos . " " . name . "`n"
if (name2>"") {
  ypos+=40,out.="TEXT 2 240 " . ypos . " " . name2 . "`n"
}
ypos+=40,out.="TEXT 2 240 " . ypos . " " . strasse . "`n"
ypos+=40,out.="TEXT 2 240 " . ypos . " " . plz . " " . ort . "`n"
if (land>"") {
  ypos+=40,out.="TEXT 2 240 " . ypos . " " . land . "`n"
}
out.="END"
msgbox,1,Adressetikett drucken,Cognitive-Drucker CL422`n`n%out%`n`nvoreingestellter Druckername: %aPrinter%
ifmsgbox,Ok
  Print(aPrinter,ansi2ascii(out))
return

filewrite(filename,byref data) {
file=%filename%.tmp
filedelete,%file%
fileappend,%data%,%file%
filemove,%filename%,%filename%.old,1
ifexist,%filename%
  exit
else
  filemove,%file%,%filename%,1
ifexist,%filename%
  filedelete,%filename%.old
else
  exit
}

WM_MOUSEMOVE(wParam,lParam) 
{ 
Global hCurs 
ifwinactive,Wählen
  return
static CurrControl, PrevControl, _TT
MouseGetPos,,,,ctrl
stringtrimleft,nmr,ctrl,6 
if (WinActive("Adressen") and nmr>3 and nmr<20) or (WinActive("Adressdaten") and nmr>12 and nmr<16)
  DllCall("SetCursor","UInt",hCurs) 
CurrControl := A_GuiControl
If (CurrControl <> PrevControl and not InStr(CurrControl, " ")) {
  ToolTip
  SetTimer,ShowHint,600
  PrevControl:=CurrControl
}
Return 

ShowHint:
SetTimer,ShowHint,Off
ToolTip % %CurrControl%_TT
SetTimer,ToolTipOff,3000
return

ToolTipOff:
ToolTip
SetTimer,ToolTipOff,Off
return
}

ansi2ascii(Zeile) {
stringreplace,zeile,zeile,Ä,Ž,a
stringreplace,zeile,zeile,Ö,™,a
stringreplace,zeile,zeile,Ü,š,a
stringreplace,zeile,zeile,ä,„,a
stringreplace,zeile,zeile,ö,”,a
stringreplace,zeile,zeile,ü,,a
stringreplace,zeile,zeile,ß,á,a
stringreplace,zeile,zeile,é,‚,a
return zeile
}

Print(byref aPrinter,DocText) {
OrgPrinter:=GetDefaultPrinter()
if ! SetDefaultPrinter(aPrinter) {
  msgbox,16,Drucken,Der Drucker %aPrinter% existiert nicht!
  return
}
VarSetCapacity(pd,66,0),NumPut(66,pd),DocName:="AHK Doc"
NumPut( aPrinter="" ? (PD_RETURNDC:=0x100) : (PD_RETURNDC:=0x100)|(PD_RETURNDEFAULT:=0x400),pd,20)
if DllCall("comdlg32\PrintDlgA","uint",&pd) {
  if (hDevMode := NumGet(pd,8)) 
    DllCall("GlobalFree","uint",hDevMode)
  if (hDevNames := NumGet(pd,12))
    DllCall("GlobalFree","uint",hDevNames)
  if (hDC := NumGet(PD, 16)) {
    PhysWidth   := DllCall("GetDeviceCaps", "UInt", hDC, "Int", 0x6E, "Int") 
    PhysHeight  := DllCall("GetDeviceCaps", "UInt", hDC, "Int", 0x6F, "Int")
    PhysOffsetX := DllCall("GetDeviceCaps", "UInt", hDC, "Int", 0x70, "Int")
    PhysOffsetY := DllCall("GetDeviceCaps", "UInt", hDC, "Int", 0x71, "Int")
    VarSetCapacity(RECT,16,0),NumPut(PhysOffsetX, RECT, 0, "Int"),NumPut(PhysOffsetY, RECT, 4, "Int")
    NumPut(PhysWidth - PhysOffsetX, RECT,  8, "Int"),NumPut(PhysHeight - PhysOffsetY, RECT, 12, "Int")
    NumPut(VarSetCapacity(DI, 20, 0), DI),NumPut(&DocName, DI, 4)
    if (DllCall("gdi32.dll\StartDoc", "UInt", hDC, "UInt", &DI, "Int") > 0) {
      if (DllCall("gdi32.dll\StartPage", "UInt", hDC, "Int") > 0) {
        DllCall("user32.dll\DrawText", "UInt", hDC, "UInt", &DocText, "Int", -1, "UInt", &RECT, "UInt", (0x0|0x200|0x10))
        DllCall("gdi32.dll\EndPage", "UInt", hDC, "Int")
      }
      DllCall("gdi32.dll\EndDoc", "UInt", hDC)
    }
    DllCall("gdi32.dll\DeleteDC", "UInt", hDC)
  }
}
SetDefaultPrinter(OrgPrinter)
}

GetDefaultPrinter() {
nSize := VarSetCapacity(sPrinter, 256)
DllCall("winspool.drv\GetDefaultPrinterA", "str", sPrinter, "UintP", nSize)
Return sPrinter
}

SetDefaultPrinter(byref drucker) {
return % DllCall("winspool.drv\SetDefaultPrinterA", "str", drucker)
}

ende:
if errorlevel {
  ListLines
  msgbox Fehler:%errorlevel%
}
exitapp

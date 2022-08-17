; https://www.autohotkey.com/boards/viewforum.php?f=10
; https://github.com/cscode2000/AdressAHK
; v1.0
; cssettings.ini für Einstellungen
; Fensterpositionen speichern
; Fenster Resize
; Umschlagdruck
; keine "letzte Änderung" wenn nur Gruppe/Land geändert wurden
; copy/mail/druck: noch nicht gespeicherte Eingaben verwenden
; Benennung Row->Col, Name...->NCol
; XP bis Win10 Anzeige vereinheitlicht
; diverse Optimierungen
;
; v0.1:
; Check auf Veränderungen vor Speichern/Mehrnutzerfähigkeit
; mit Land + letzter Änderung
; statt Buttons: Icons + Tooltips 
; variable Zahl Telefonnummern/Mailadressen
; zum Wählen der Tel.nr. mit Fritzbox fb_anruf.ahk aufrufen
; übergabe an Mailprogramm/Zwischenablage
; direkte Etikettendruckeransteuerung (Nur-Text)
#singleinstance force 
#NoEnv
Menu,Tray,Icon,pifmgr.dll,16
Menu,Tray,Add,&Anzeigen,search
Menu,Tray,Default,&Anzeigen
settitlematchmode,Regex
setworkingdir,%A_ScriptDir%
adr_ini=%a_appdata%\cssettings.ini
database=%a_appdata%\adress.dat
if 1>
  database=%1%
iniread,guipos,%adr_ini%,adressen,guipos,%a_space%
iniread,gui2pos,%adr_ini%,adressen,gui2pos,%a_space%
iniread,gui3pos,%adr_ini%,adressen,gui3pos,%a_space%
iniread,gui4pos,%adr_ini%,adressen,gui4pos,%a_space%
iniread,groups,%adr_ini%,adressen,groups,%a_space%
if groups=
  iniwrite,% groups:="Firma|Privat|obsolet|alle",%adr_ini%,adressen,groups
sysget,cpt,4
sysget,bdrx,32
sysget,bdry,33
onexit,ende
hCurs:=DllCall("LoadCursor","UInt",NULL,"Int",32649,"UInt") 
OnMessage(0x200,"WM_MOUSEMOVE") 
regread,Col1,HKCU,Software\Clients\Mail
regread,Col2,HKLM,SOFTWARE\Clients\Mail\%Col1%\Protocols\mailto\shell\open\command
regexmatch(Col2,"\x22(.+?)\x22",Col)
if strlen(Col1)<7
  regexmatch(Col2,"(.+?)\s",Col)
Mailer:=Col1

suchtext_TT=Namensbestandteil suchen
group_TT=Adressen-Gruppe wählen (Alt+g)
ButtonBearbeiten_TT=Adressdaten bearbeiten (Enter)
ButtonNeu_TT=Neuen Eintrag erstellen (Einfg)
ButtonLöschen_TT=Eintrag löschen (Entf)
ButtonCopy_TT=Adresse in Zwischenablage kopieren (F4)
Anrufen_TT=Telefonnummer anrufen (F5)
EMailen_TT=E-Mail schreiben (F6)
drk2_TT=Briefumschlag bedrucken (F8)
ButtonDruck_TT=Adressetikett drucken (F9)
privat_TT=Aktiviert = privater Absender wird gedruckt`nAusgegraut = kein Absender
ButtonSettings_TT=Voreinstellungen öffnen
Anrufen2_TT:=Anrufen_TT
email2_TT:=EMailen_TT
ButtonCopy2_TT:=ButtonCopy_TT

Gui, +Owner +Resize +LastFound
WinGet, MainID, ID
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
Gui,Add,Picture,ys-3 w20 h-1 vdrk2 gButtonDruck Icon157,c:\windows\system32\shell32.dll
Gui,Add,Picture,ys-3 w20 h-1 vButtonDruck gButtonDruck Icon137,c:\windows\system32\shell32.dll
Gui,Add,Picture,ys-3 w20 h-1 vButtonSettings gButtonSettings Icon166,c:\windows\system32\shell32.dll
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
  stringreplace,Col,A_Loopfield,¶,`n,a
  StringSplit,Col,Col,|
  if instr(Col1,suchtext) and (instr(group,Col12)=1 or group="alle")
    LV_Add("", Col1, Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9,Col10,Col11,Col12,Col13,Col14)
}
GuiControl,Show,MyListView
blockinput,off
Gui,Show,%guipos%,% "Adressen (" . LV_GetCount() . ")"
guipos=
LV_Modify(1,"Focus Select")
send, {Ctrl up} 
send, {Shift up}
return

GuiEscape:
GuiClose:
ExitApp

GuiSize:
WinGetPos, winX, winY, winW, winH, ahk_id %MainID%
GuiControl, Move, MyListView, % "w" winW-6-bdrx "h" winH-52-2*bdry
return

#ifwinactive,^Wählen$
~right::
~left::
~up::
~down::
CursorTaste=1
return

#IfWinActive,^Adress(dat)?en,&Gruppe:
<^>!<::return
F4::goto,ButtonCopy
F5::goto,ButtonAnrufen
F6::goto,ButtonEMail
F8::goto,ButtonDruck2
F9::goto,ButtonDruck

#ifwinactive,^Adressen
Del::
GuiControlGet, Focus, Focus
if (Focus="SysListView321")
  goto,ButtonLöschen
return
~down::
GuiControlGet, Focus, Focus
if (Focus="Edit1")
  send,{tab}
return

*Ins::
ButtonNeu:
neuanlage=1

ButtonLöschen:
neuanlage:=(neuanlage=1) ? neuanlage : -1

ButtonBearbeiten:
MyListView:
if (A_GuiEvent="ColClick")
  return
neuanlage*=2
ifwinexist,ahk_id %EditID% 
{
  msgbox,,Achtung,Erst anderes Adressdaten-Fenster schließen!
  winactivate
  return
}
if neuanlage=2
  Loop % LV_GetCount("Col")
    Col%a_index%:=""
else {
  ControlGet, Zeile, List, Count Focused, SysListView321
  ControlGet, Zeile2, List, Focused, SysListView321
  stringreplace,zeile2,zeile2,%A_Tab%,|,a
  stringreplace,zeile2,zeile2,`n,¶,a
  fileread,f1,%database%
  Loop,parse,f1,`n,`r
    if (A_Loopfield=zeile2) {
      stringreplace,zeile2,zeile2,¶,`n,a
      stringsplit,Col,zeile2,|
      zeile2=
      break
    }
}
if (zeile2>"") {
  msgbox,1,Adressdaten,Der Datensatz wurde zwischenzeitlich geändert und kann nicht bearbeitet werden.`n`nMöchten Sie die Datei neu einlesen?
ifmsgbox,Ok
  run,%A_AhkPath% "%A_ScriptFullPath%"
return
}
if (neuanlage<>-2) {
  Gui,2:Font,s11
  EOpts:="x100 yp-3 W300",TOpts:="x12 yp+30"
  Gui,2:Add,Text,%TOpts% y10,&Name:
  Gui,2:Add,Edit,%EOpts% vNCol1,%Col1%
  Gui,2:Add,Text,%TOpts%,Na&me 2:
  Gui,2:Add,Edit,%EOpts% vNCol2,%Col2%
  Gui,2:Add,Text,%TOpts%,St&raße:
  Gui,2:Add,Edit,%EOpts% vNCol3,%Col3%
  Gui,2:Add,Text,%TOpts%,&PLZ / Ort:
  Gui,2:Add,Edit,x100 yp-3 W50 vNCol4,%Col4%
  Gui,2:Add,Edit,x155 yp W245 vNCol5,%Col5%
  Gui,2:Add,Text,%TOpts%,&Telefon:
  Gui,2:Add,Edit,%EOpts% R3 vNCol6,%Col6%
  Gui,2:Add,Text,x12 yp+62,&E-Mail:
  Gui,2:Add,Edit,%EOpts% R2 vNCol7,%Col7%
  Gui,2:Add,Text,x12 yp+46,Fa&x:
  Gui,2:Add,Edit,%EOpts% vNCol8,%Col8%
  Gui,2:Add,Text,%TOpts%,&Internet:
  Gui,2:Add,Edit,%EOpts% -Wrap -VScroll vNCol9,%Col9%
  Gui,2:Add,Text,%TOpts%,&Kdnr:
  Gui,2:Add,Edit,%EOpts% vNCol10,%Col10%
  Gui,2:Add,Text,%TOpts%,Bemerk&ung:
  Gui,2:Add,Edit,%EOpts% R4 vNCol11,%Col11%
  Gui,2:Add,Text,x12 yp+78 vnx3,&Land:
  Gui,2:Add,Edit,x100 yp-3 W120 vNCol13,%Col13%
  Gui,2:Add,Text,x230 yp+3 vnx4,&Gruppe:
  Gui,2:Add,ComboBox,x285 yp-3 W115 vNCol12,%Col12%||%groups%
  Gui,2:Font,s8
  Gui,2:Add,Button,x410 y8 W80 vdrk3 gButtonDruck,Brie&fumschlag
  Gui,2:Add,Button,x410 y36 W80 vdrk gButtonDruck,Etiketten&druck
  Gui,2:Add,Button,x410 y64 W80 vbüwa gButtonDruck,Etikett Bü&Wa
  Gui,2:Add,Checkbox,x417 y91 vprivat Check3,Pri&vatabs.

  Gui,2:Add,Picture,x430 y129 w32 h-1 vAnrufen2 gButtonAnrufen Icon197,c:\windows\system32\shell32.dll
  if Mailer>
    Gui,2:Add,Picture,x430 y179 w32 h-1 vemail2 gButtonEMail,%Mailer%
  else
    Gui,2:Add,Picture,x430 y179 w32 h-1 vemail2 gButtonEMail Icon215,c:\windows\system32\shell32.dll
  Gui,2:Add,Picture,x436 y240 w20 h-1 vButtonCopy2 gButtonCopy Icon55,c:\windows\system32\shell32.dll
 
  Col14n:=substr(Col14,5,2) . "." . substr(Col14,3,2) . "." . substr(Col14,1,2) . " " . substr(Col14,7,2) . ":" . substr(Col14,9,2)
  Gui,2:Add,Text,x410 y300 0x2 vnx1,letzte Änderung:`n%Col14n%
  Gui,2:Add,Button,% "Default x410 W80 vSpeichern y" 373-cpt-2*bdry,&Speichern
  Gui,2:Add,Button,x410 yp+28 W80 vnx2,A&bbrechen
  Gui,2: +LastFound +Resize
  WinGet, EditID, ID
  Gui,2:Show,%gui2pos%,Adressdaten
  return
}

2ButtonSpeichern:
gosub,checkchanged 
if (neuanlage=-2) {
  msgbox,257,Adressdaten,Dieser Datensatz wird unwiderruflich gelöscht:`n`n%alt%
  ifmsgbox,OK
  ifwinexist,ahk_id %EditID%
  {
    msgbox,,Achtung,Erst anderes Adressdaten-Fenster schließen!
    winactivate
  }
  else
  {
    cflag=1
    hnd:=fileopen(database,"r-w")
    fileread,f1,%database% 
    Loop,parse,f1,`n,`r
      f2.=(A_Loopfield=alt . "|" . Col14) ? cflag:="" : A_Loopfield . "`n"
    stringtrimright,f1,f2,1
    fileclose(hnd)
    if !cflag {
      if filewrite(database,f1) {
        LV_Delete(zeile)
        Gui,Show,,% "Adressen (" . LV_GetCount() . ")"
      } else
        msgbox,,Adressdaten,Vorgang nicht möglich`, da die Daten zur Zeit gesperrt sind.
    }
    else
    {
      msgbox,1,Adressdaten,Vorgang nicht möglich`, da die Daten inzwischen geändert wurden.`n`nMöchten Sie die Daten neu einlesen?
      ifmsgbox,Ok
        run,%A_AhkPath% "%A_ScriptFullPath%"
    }
  }
} else {
  if (alt=neu) or (neuanlage=2 and neu="||")
  {
    gosub,savepos_destroy2
    return
  }
  if (FocusedControl<>"Speichern") 
    msgbox,35,Adressdaten,Geänderte Daten speichern?
  ifmsgbox,Cancel
    return
  ifmsgbox,No
  {
    gosub,savepos_destroy2
    return
  }
  formattime,Col14n,,yyMMddHHmm
  if (alt<>neu) {
    cflag=
    if neuanlage=2
      f1.=(f1="" ? "" : "`n") . neu . "|" . Col14n
    else {
      hnd:=fileopen(database,"r-w")
      fileread,f1,%database% 
      Loop,parse,f1,`n,`r
        f2.=(A_Loopfield=(alt . "|" . Col14)) ? alt_basic=neu_basic ? neu . "|" . (cflag:=Col14) . "`n" : neu . "|" . (cflag:=Col14n) . "`n" : A_Loopfield . "`n"
      stringtrimright,f1,f2,1 
fileclose(hnd)
    }
    if (neuanlage=2) and !fileexist(database)
      fileappend,#,%database% 
    if (neuanlage=2) or cflag {
      if filewrite(database,f1) {
        settimer,list_akt,-1
      } else
        msgbox,,Adressdaten,Speichern nicht möglich`, da die Daten zur Zeit gesperrt sind.
    }
    else 
    {
      msgbox,1,Adressdaten,Die Daten können nicht gespeichert werden`, da sie inzwischen geändert wurden.`n`nMöchten Sie die Daten neu einlesen?
      ifmsgbox,Ok
      {
        onexit
        run,%A_AhkPath% "%A_ScriptFullPath%"
      }
    }
  }
}
return

2ButtonAbbrechen:
2GuiEscape:
2GuiClose:
gosub,checkchanged
if (alt=neu) or (DllCall("MessageBox", "Int", "0", "Str", "Verlassen und Änderungen verwerfen?", "Str", "Adressdaten", "Int", 35)=6)
  gosub,savepos_destroy2
return


2GuiSize:
WinGetPos, winX, winY, winW, winH, ahk_id %EditID%
winw:= winw<511 ? 511 : winw
winh:= winh<432 ? 432 : winh
loop,11
  if a_index<>4
    GuiControl, Move, NCol%a_index%, % "w" winW-210
GuiControl, Move, NCol5, % "w" winW-265
GuiControl, Move, drk3, % "x" winW-100
GuiControl, Move, drk, % "x" winW-100
GuiControl, Move, büwa, % "x" winW-100
GuiControl, Move, privat, % "x" winW-93
GuiControl, Move, ButtonCopy2, % "x" winW-74
GuiControl, Move, Anrufen2, % "x" winW-80
GuiControl, Move, email2, % "x" winW-80
GuiControl, Move, nx1, % "x" winW-100
GuiControl, Move, nx2, % "x" winW-100
GuiControl, Move, Speichern, % "x" winW-100
GuiControl, Move, NCol12, % "w" winW-395 " y" winh-28-cpt-2*bdry
GuiControl, Move, NCol13, % "y" winh-28-cpt-2*bdry
GuiControl, Move, nx3, % "y" winh-25-cpt-2*bdry
GuiControl, Move, nx4, % "y" winh-25-cpt-2*bdry
GuiControl, Move, NCol11, % "h" winh-329-cpt-2*bdry
return

list_akt:
gosub,savepos_destroy2
stringreplace,NCol6,NCol6,¶,`n,a
stringreplace,NCol7,NCol7,¶,`n,a
stringreplace,NCol11,NCol11,¶,`n,a
if instr(NCol1,suchtext) and (instr(group,NCol12)=1 or group="alle")
  if neuanlage=2
    LV_Add("Focus Select Vis",NCol1,NCol2,NCol3,NCol4,NCol5,NCol6,NCol7,NCol8,NCol9,NCol10,NCol11,NCol12,NCol13,Col14n)
  else
    LV_Modify(Zeile,"Focus Select",NCol1,NCol2,NCol3,NCol4,NCol5,NCol6,NCol7,NCol8,NCol9,NCol10,NCol11,NCol12,NCol13,cflag)
else
  if neuanlage<>2 
    LV_Delete(zeile)
neuanlage=
Gui,Show,,% "Adressen (" . LV_GetCount() . ")"
return

checkchanged: 
Gui,2:Submit,NoHide
GuiControlGet, FocusedControl, FocusV
alt_basic:="",neu_basic:="",f2:=""
loop,13
{
  stringreplace,NCol%a_index%,NCol%a_index%,`n,¶,a 
  stringreplace,NCol%a_index%,NCol%a_index%,|,,a
}
loop,10
{
  alt_basic.= Col%a_index% . "|"
  neu_basic.=NCol%A_Index% . "|"
}
alt_basic.=Col11
alt=%alt_basic%|%Col12%|%Col13%
stringreplace,alt,alt,`n,¶,a
neu_basic.=NCol11
neu=%neu_basic%|%NCol12%|%NCol13%
return

ButtonCopy:
if winactive("ahk_id " MainID) {
  ControlGet, Zeile, List, Focused, SysListView321
  stringsplit,Col,zeile,%A_Tab%
} else
  N=N
Gui,2:Submit,NoHide
clipboard:=%N%Col1 . (%N%Col2>"" ? "`r`n" . %N%Col2 : "") . "`r`n" . %N%Col3 . "`r`n"
   . %N%Col4 . " " . %N%Col5 . "`r`n" . (%N%Col10>"" ? "Kd.nr.: " . %N%Col10 : "")
tooltip,=== OK ===
sleep,900
tooltip
return

ButtonAnrufen:
ifwinexist,^Wählen$
{
  winactivate
  return
}
CursorTaste=0
Gui,2:Submit,NoHide
if winactive("ahk_id " MainID)
  ControlGet, NCol6, List, Focused Col6, SysListView321
var1:=NCol6
nr_anz:=1
regexmatch(var1,"([^\r\n¶]+)",tel)
Gui,3:Add,Text,,Welche Nummer soll gewählt werden?`n
Gui,3:Add,Radio,x20 gNrSelect vtel Checked,%tel1%
loop,parse,var1,`n¶,`r
  if (a_index>1) {
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
  Gui,3:Show,%gui3pos%,Wählen
  return
}

NrSelect:
if CursorTaste=1
  return 

3ButtonOK:
Gui,3:Submit

winGetPos, X, Y, , ,Wählen
if x>
  gui3pos=x%x%y%y%
Gui,3:Destroy

nr:=regexreplace(tel%tel%,"(\D*)([\d- /\*#]+)(.*)","$2")
if nr>
{
  Gui,4:Add,Text,Section x12 y20,Nummer: 
  Gui,4:Add,Edit,xs ys W250 -WantReturn H20 vnr,%nr%
  Gui,4:Add,Checkbox,gclip_jn ys+3,Nummer übermitteln
  Gui,4:Add,Button,Default Section x12 W80,Ok
  Gui,4:Add,Button,xp+90 ys W80,&Abbrechen
  ControlGet, NCol1, List, Focused Col1, SysListView321
  Gui,4:Show,%gui4pos%,Telefonanruf %NCol1%
  return

  4ButtonOk:
  Gui,4:Submit
  if nr>
    run,%A_AhkPath% %a_scriptdir%\fb_anruf.ahk "%clip%%nr%",,min

  4ButtonAbbrechen:
  4GuiEscape:
  4GuiClose:
  winGetPos, X, Y, , ,Telefonanruf
  gui4pos=x%x%y%y%
  Gui,4:Destroy
  return

  clip_jn:
  clip:=clip="" ? "#31#," : ""
  return
}
return
3ButtonAbbrechen:
3GuiEscape:
winGetPos, X, Y, , ,Wählen
gui3pos=x%x%y%y%
Gui,3:Destroy
return

ButtonEMail:
Gui,2:Submit,NoHide
if winactive("ahk_id " MainID)
  ControlGet, NCol7, List, Focused Col7, SysListView321
var1:=NCol7
regexmatch(var1,"([\w\.\-\+_]+@[\w\.\-]+)",var)
var1:=var1="" ? NCol7 : var1
run,%Mailer% mailto:%var1%
return

ButtonDruck2:
out=2
ButtonDruck:
if winactive("ahk_id " MainID) {
  ControlGet, Zeile, List, Focused, SysListView321
  stringsplit,Col,zeile,%A_Tab%
} else
  N=N
Gui,2:Submit,NoHide
iniread,absender,%adr_ini%,adressen,absender,%a_space%
iniread,abs_privat,%adr_ini%,adressen,abs_privat,%a_space%
if absender=
  iniwrite,% absender:="Absender Geschäft...",%adr_ini%,adressen,absender
if abs_privat=
  iniwrite,% abs_privat:="Absender privat...",%adr_ini%,adressen,abs_privat

if (a_guicontrol="drk2" or a_guicontrol="drk3" or out=2) {
  out:=%N%Col1 . "`n"
  if (%N%Col2>"")
    out.=%N%Col2 . "`n"
  out.=%N%Col3 . "`n"
  out.=%N%Col4 . " " . %N%Col5 . "`n"
  if (Col13>"")
    out.=%N%Col13 . "`n"

  msgbox,1,Briefumschlag bedrucken,%out%
  ifmsgbox,Ok
    Print2(out)
} else {
  out:="!A`n"
  if privat<>-1
    out.="TEXT 1 240 0 " . (privat=1 ? abs_privat : absender) . "`n"
  ypos:=(privat=-1) ? -40 : 10
  if (a_guicontrol="büwa") {
    ypos+=40,out.="TEXT 2 240 " . ypos . " BÜWA`n"
  }
  ypos+=40,out.="TEXT 2 240 " . ypos . " " . %N%Col1 . "`n"
  if (%N%Col2>"") {
    ypos+=40,out.="TEXT 2 240 " . ypos . " " . %N%Col2 . "`n"
  }
  ypos+=40,out.="TEXT 2 240 " . ypos . " " . %N%Col3 . "`n"
  ypos+=40,out.="TEXT 2 240 " . ypos . " " . %N%Col4 . " " . %N%Col5 . "`n"
  if (Col13>"") {
    ypos+=40,out.="TEXT 2 240 " . ypos . " " . %N%Col13 . "`n"
  }
  out.="QUANTITY 1`nEND"
  msgbox,1,Adressetikett drucken,Modell Cognitive CL422`n`n%out%
  ifmsgbox,Ok
    Print(ansi2ascii(out))
}
n=
return

ButtonSettings:
run,"%adr_ini%"
return

filewrite(filename,byref data) {
filecopy,%filename%,%filename%.bak,1
if fileexist(filename . ".bak") and (hnd:=fileopen(filename,"w-w"))>0 {
  ok:=DllCall("WriteFile","UInt",hnd,"UChar",&data,"UInt",strlen(data),"UInt *", Written,"UInt",0)
  fileclose(hnd)
  if strlen(data)=written and ok
    return 1
  filecopy,%filename%.bak,%filename%,1
  msgbox,,Fehler %ok% beim Speichern von %filename%,% strlen(data) . " Bytes gesamt, gespeichert: " . written
}
}

fileopen(filename,flags) {
Access= 0
GENERIC_WRITE= 0x40000000
GENERIC_READ = 0x80000000
FILE_SHARE_READ  = 1
CREATE_ALWAYS    = 2
OPEN_ALWAYS      = 4
Share:=FILE_SHARE_READ
if (flags="w-w") {
  Access:=GENERIC_WRITE
  Creation:=CREATE_ALWAYS
}
if (flags="r-w") {
  Access:=GENERIC_READ
  Creation:=OPEN_ALWAYS
}
hFile := DllCall("CreateFile", "Str", FileName, "UInt", Access, "UInt", Share, "UInt", 0, "UInt", Creation, "UInt", 0, "UInt", 0)
if hFile<0
  msgbox,,Fehler %hFile% in fileopen("%filename%"`,"%flags%"),Die Datei kann nicht gespeichert werden.
return % hFile
}

fileclose(handle) {
DllCall("CloseHandle", "UInt", handle) 
}

WM_MOUSEMOVE(wParam,lParam) 
{ 
Global hCurs,MainID,EditID
if A_Gui>2
  return
static CurrControl, PrevControl, _TT
MouseGetPos,,,,ctrl
stringtrimleft,nmr,ctrl,6
if (WinActive("ahk_id " MainID) and nmr>3 and nmr<20) or (WinActive("ahk_id " EditID) and nmr>12 and nmr<16)
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

Print(DocText) {
global adr_ini
iniread,aPrinter,%adr_ini%,adressen,EtikettP,%a_space%
if aPrinter=
  iniwrite,% aPrinter:="Adressetiketten",%adr_ini%,adressen,EtikettP
if DllCall("winspool.drv\OpenPrinterA", "str", aPrinter, "UInt *", hDC, "UInt",0) {
	DocType:="RAW",DocName:="AHK Doc"
	VarSetCapacity(DocInfo, 12, 0),NumPut(&DocName, DocInfo,0,"UInt"),NumPut(&DocType, DocInfo, 8, "UInt")
	if DllCall("winspool.drv\StartDocPrinterA", "UInt", hDC, "uint", 1, "UInt", &DocInfo) {
    if DllCall("winspool.drv\StartPagePrinter", "UInt", hDC) {
       DllCall("winspool.drv\WritePrinter", "UInt", hDC, "UInt", &DocText, "uint", strlen(DocText), "UInt *", Written)
       DllCall("winspool.drv\EndPagePrinter", "UInt", hDC)
    }
    DllCall("winspool.drv\EndDocPrinter", "UInt", hDC)
  }
  DllCall("winspool.drv\ClosePrinter", "UInt", hDC)
}
}

Print2(DocText) {
global adr_ini,absender,abs_privat,privat
iniread,aPrinter,%adr_ini%,adressen,UmschlagP,%a_space%
if aPrinter=
  iniwrite,%a_space%,adr_ini,adressen,UmschlagP
iniread,xpos,%adr_ini%,adressen,Rand_L,193
iniread,ypos,%adr_ini%,adressen,Rand_O,105
VarSetCapacity(pDevModeOutput , DllCall("Winspool.drv\DocumentPropertiesA", UInt, MainhWnd, UInt, pPrinter, UInt, &aPrinter, UInt, 0, UInt, 0, UInt, 0), 0)
if DllCall("Winspool.drv\DocumentPropertiesA", UInt, MainhWnd, UInt, pPrinter, UInt, &aPrinter, UInt, &pDevModeOutput, UInt, 0, UInt, 2) > 0 {
  NumPut(2, pDevModeOutput, 44, "Short") ; 1=hoch 2=quer
  NumPut(9, pDevModeOutput, 46, "Short") ; A4 210 x 297 mm 
  NumPut(1, pDevModeOutput, 54, "Short") ; Anzahl Kopien
  NumPut(4, pDevModeOutput, 56, "Short") ; Druckerschacht	UPPER = 1, MANUAL = 4, ENVELOPE = 5, AUTO = 7
  NumPut(1, pDevModeOutput, 60, "Short")
  if !DllCall("GetModuleHandle", "str", "gdiplus")
    DllCall("LoadLibrary", "str", "gdiplus")
  VarSetCapacity(si, 16, 0), si := Chr(1)
  DllCall("gdiplus\GdiplusStartup", "uint*", pToken, "uint", &si, "uint", 0)
    hDc :=  DllCall("Gdi32.dll\CreateDC", "Str", "", UInt, &aPrinter, "Str", "", UInt, &pDevModeOutput)

  global PhysWidth,PhysHeight,HDC_ydpi
  PhysWidth   := DllCall("GetDeviceCaps", "UInt", hDC, "Int", 0x6E) 
  PhysHeight  := DllCall("GetDeviceCaps", "UInt", hDC, "Int", 0x6F)
  PhysOffsetX := DllCall("GetDeviceCaps", "UInt", hDC, "Int", 0x70)
  PhysOffsetY := DllCall("GetDeviceCaps", "UInt", hDC, "Int", 0x71)
 	HDC_ydpi    := DllCall("GetDeviceCaps", "UInt", hDC, "int", 0x5A)
  DocName:="AHK Doc",NumPut(VarSetCapacity(DI, 20, 0), DI),NumPut(&DocName, DI, 4)
  if (DllCall("gdi32.dll\StartDoc", "UInt", hDC, "UInt", &DI, "Int") > 0) {
    if (DllCall("gdi32.dll\StartPage", "UInt", hDC) > 0) {
      DllCall("gdiplus\GdipCreateFromHDC", "uint", hdc, "uint*", G)
      DllCall("gdiplus\GdipSetPageUnit",UInt,G,"int",2) 
      xpos:=xpos*PhysWidth/297,ypos:=ypos*PhysHeight/210
      if privat<>-1
        DrawString(G, privat=1 ? abs_privat : absender, xpos, ypos, PhysWidth-xpos, PhysHeight-ypos, "Arial", 0, 0, 9)
      ypos:=111*PhysHeight/210
      DrawString(G, DocText, xpos, ypos, PhysWidth-xpos, PhysHeight-ypos)
      DllCall("gdi32.dll\EndPage", "UInt", hDC)
    }
    DllCall("gdi32.dll\EndDoc", "UInt", hDC)
  }
  DllCall("gdi32.dll\DeleteDC", "UInt", hDC)
  DllCall("gdiplus\GdipDeleteGraphics", "uint", G)
  DllCall("gdiplus\GdiplusShutdown", "uint", pToken)
  if hModule := DllCall("GetModuleHandle", "str", "gdiplus")
    DllCall("FreeLibrary", "uint", hModule)
} else
  msgbox,,Drucken,Fehler bei Drucker %aPrinter%
}

DrawString(G, DocText, xpos = 0, ypos = 0, Width = "full", Height = "full", Font = "Arial", Style = 0, Align = 0, sizeinpoints = 12)
; Style = Bold := 1 , Italic := 2 , Underline := 4 , Strikeout := 8
; sizeinpoints = font size in points(1 Point = ydpi / 72), wie in Textverarbeitung
{
  global PhysWidth,PhysHeight,HDC_ydpi
	size := Round((HDC_ydpi / 72) * sizeinpoints)
  Width:= Width>":" ? PhysWidth : Width
  Height:= Height>":" ? PhysHeight : Height
  VarSetCapacity(RC, 16)
  NumPut(xpos, RC, 0, "float"), NumPut(ypos, RC, 4, "float"), NumPut(width, RC, 8, "float"), NumPut(height, RC, 12, "float")
		
  nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, "uint", &Font, "int", -1, "uint", 0, "int", 0)
  VarSetCapacity(wFont, nSize*2)
  DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, "uint", &Font, "int", -1, "uint", &wFont, "int", nSize)
  DllCall("gdiplus\GdipCreateFontFamilyFromName", "uint", &wFont, "uint", 0, "uint*", hFamily)
  DllCall("gdiplus\GdipCreateFont", "uint", hFamily, "float", Size, "int", Style, "int", 0, "uint*", hFont)
  DllCall("gdiplus\GdipCreateStringFormat", "int", 0x4000, "int", 0, "uint*", hFormat)
  DllCall("gdiplus\GdipCreateSolidFill", "int", 0xff000000, "uint*", pBrush)
  DllCall("gdiplus\GdipSetStringFormatAlign", "uint", hFormat, "int", Align)
  DllCall("gdiplus\GdipSetTextRenderingHint", "uint", G, "int", 3)
  DllCall("gdiplus\GdipMeasureString", UInt, G, "wstr", DocText, "int", -1, UInt, hFont, UInt, &RC, UInt, hFormat)
  nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, "uint", &DocText, "int", -1, "uint", 0, "int", 0)
  VarSetCapacity(wString, nSize*2)
  DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, "uint", &DocText, "int", -1, "uint", &wString, "int", nSize)
  DllCall("gdiplus\GdipDrawString", "uint", G, "uint", &wString, "int", -1, "uint", hFont, "uint", &RC, "uint", hFormat, "uint", pBrush)
  DllCall("gdiplus\GdipDeleteBrush", "uint", pBrush)
  DllCall("gdiplus\GdipDeleteStringFormat", "uint", hFormat)
  DllCall("gdiplus\GdipDeleteFont", "uint", hFont)
  DllCall("gdiplus\GdipDeleteFontFamily", "uint", hFamily)
}	

savepos_destroy2:
winGetPos,winX,winY,winw,winh,ahk_id %EditID%
winw:= winw<511 ? 511 : winw-bdrx-bdry
winh:= winh<(378+cpt+2*bdry) ? 420 : winh-cpt-bdrx-bdry
gui2pos=x%winx%y%winy%w%winw%h%winh%
Gui,2:Destroy
EditID=
return

ende:
winGetPos,X,Y,W,H,ahk_id %MainID%
w:= w<627 ? 627 : w-2*bdrx
h:= h<500 ? 500 : h-cpt-2*bdry
iniwrite,x%x%y%y%w%w%h%h%,%adr_ini%,Adressen,guipos
ifwinexist,ahk_id %EditID%
{
  winactivate
  gosub,2ButtonSpeichern
}
if gui2pos>
  iniwrite,%gui2pos%,%adr_ini%,Adressen,gui2pos
winGetPos, X, Y, , ,Wählen
if x>
  iniwrite,x%x%y%y%,%adr_ini%,Adressen,gui3pos
else
  if gui3pos>
    iniwrite,%gui3pos%,%adr_ini%,Adressen,gui3pos
winGetPos, X, Y, , ,Telefonanruf
if x>
  iniwrite,x%x%y%y%,%adr_ini%,Adressen,gui4pos
else
  if gui4pos>
    iniwrite,%gui4pos%,%adr_ini%,Adressen,gui4pos
exitapp

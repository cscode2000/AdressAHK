#singleinstance force 
settitlematchmode,2
setworkingdir,%A_ScriptDir%
if 1=
  database=%appdata%\adress.dat
onexit,ende
Gui,Add,Text,Section x12 y10,Suche: 
Gui,Add,Edit,ys-2 W100 H20 gsearch vsuchtext,%suchtext%
Gui,Font,s11
Gui,Add,ListView,x12 r30 w800 gMyListView vMyListView Sort -Multi,Name|Name 2|Straße|PLZ|Ort|Telefon|Telefon 2|Fax|Internet|Kdnr|Bemerkung|Gruppe|Land
f1=150|100|100|50|100|100|100|100|100|50|1|20|50
loop,parse,f1,|
  LV_ModifyCol(a_index,a_loopfield)
Gui,Font,s8
Gui,Add,Text,Section x170 y10,Gruppe:
Gui,Add,DropDownList,ys-2 w80 vgroup gsearch,Firma|Privat|obsolet|alle
Gui,Add,Button,ys-3 W80 Default,&Bearbeiten 
Gui,Add,Button,ys-3 W80,&Neu
Gui,Add,Button,ys-3 W80,&Löschen
Gui,Add,Button,ys-3 W80,&Anrufen
Gui,Add,Button,ys-3 xp+150 W80,B&eenden
fileread,f1,%database%
search:
Gui,Submit,NoHide
blockinput,on
LV_Delete()
GuiControl,Hide,MyListView
Loop,parse,f1,`n,`r
{
  StringSplit,row,A_Loopfield,|
  stringreplace,row11,row11,¶,`n,a
  if instr(row1,suchtext) and (instr(group,row12)=1 or group="alle")
    LV_Add("", row1, row2,row3,row4,row5,row6,row7,row8,row9,row10,row11,row12,row13)
}
GuiControl,Show,MyListView
blockinput,off
Gui,Show,,% "Adressen (" . LV_GetCount() . ")"
LV_Modify(1,"Focus Select")
send, {Ctrl up}
send, {Shift up}
return

ButtonLöschen:
neuanlage=-1

ButtonNeu:
neuanlage:=(neuanlage=-1) ? neuanlage : 1

ButtonBearbeiten:
MyListView:
neuanlage*=2 
ifwinexist,Bearbeiten,Name 2
{
  msgbox,,Achtung,Erst anderes Bearbeiten-Fenster schließen!
  winactivate
  return
}
if !(A_GuiEvent = "DoubleClick" or A_GuiEvent = "Normal")
  return
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
  Gui,2:Add,Edit,%EOpts% vTelefon,%Row6%
  Gui,2:Add,Text,%TOpts%,Te&lefon 2:
  Gui,2:Add,Edit,%EOpts% vTelefon2,%Row7%
  Gui,2:Add,Text,%TOpts%,Fa&x:
  Gui,2:Add,Edit,%EOpts% vFax,%Row8%
  Gui,2:Add,Text,%TOpts%,&Internet:
  Gui,2:Add,Edit,%EOpts% vInternet,%Row9%
  Gui,2:Add,Text,%TOpts%,&Kdnr:
  Gui,2:Add,Edit,%EOpts% vKdnr,%Row10%
  Gui,2:Add,Text,%TOpts%,Bemerk&ung:
  Gui,2:Add,Edit,%EOpts% R4 vBemerkung,%Row11%
  Gui,2:Add,Text,x12 yp+80,Lan&d:
  Gui,2:Add,Edit,x100 yp-3 W185 vLand,%Row13%
  Gui,2:Add,Text,x295 yp+3,&Gruppe:
  Gui,2:Add,Edit,x350 yp-3 W50 vGruppe,%Row12%
  Gui,2:Font,s8
  Gui,2:Add,Button,x410 y118 W80 vtel1,&Anrufen
  Gui,2:Add,Button,x410 y145 W80 vtel2,Anrufen
  Gui,2:Add,Button,x410 y172 W80,Anrufen

  Gui,2:Add,Button,Section x12 W80,A&bbrechen
  Gui,2:Add,Button,Default xp+90 ys W80,&Speichern
  Gui,2:Add,Button,xp+130 ys W80,&Etikett
  Gui,2:Add,Button,xp+90 ys W80,Etikett Bü&Wa
  Gui,2:Add,Checkbox,ys+4 vprivat,Pri&vatabs.
  Gui,2:Show,,Bearbeiten
  return
}

2ButtonSpeichern:
Gui,2:Submit,NoHide
GuiControlGet, FocusedControl, FocusV
alt:="",f2:=""
stringreplace,row11,row11,`n,¶,a
stringreplace,Bemerkung,Bemerkung,`n,¶,a
loop,12
  alt.= row%a_index% . "|"
alt.=row13
neu=%name%|%name2%|%Strasse%|%PLZ%|%Ort%|%Telefon%|%Telefon2%|%Fax%|%Internet%|%Kdnr%|%Bemerkung%|%Gruppe%|%Land%
if (neuanlage=-2) {
  msgbox,4,Bearbeiten,Datensatz löschen?`n%alt%
  ifmsgbox,No
    return
  LV_Delete(zeile)
  Gui,Show,,% "Adressen (" . LV_GetCount() . ")"
  Loop,parse,f1,`n,`r
    f2.=(A_Loopfield=alt) ? "" : A_Loopfield . "`n"
  stringtrimright,f1,f2,1
  filewrite(database,f1)
} else {
  if (alt<>neu and FocusedControl<>"&Speichern")
    msgbox,4,Bearbeiten,Geänderte Daten speichern?
  ifmsgbox,No
    return
  Gui,2:Destroy
  if (neuanlage=2 and neu="|||||||||||")
    return
  if (alt<>neu) {
    if neuanlage=2
      f1.=(f1="" ? "" : "`n") . neu
    else {
      Loop,parse,f1,`n,`r
        f2.=(A_Loopfield=alt) ? neu . "`n" : A_Loopfield . "`n"
      stringtrimright,f1,f2,1
    }
    filewrite(database,f1)
    settimer,list_akt,-1
  }
}
return

list_akt:
stringreplace,Bemerkung,Bemerkung,¶,`n,a
if instr(name,suchtext) and (instr(group,gruppe)=1 or group="alle")
  if neuanlage=2
    LV_Add("Focus Select Vis",name,name2,Strasse,PLZ,Ort,Telefon,Telefon2,Fax,Internet,Kdnr,Bemerkung,Gruppe,Land)
  else
    LV_Modify(Zeile,"Focus Select",name,name2,Strasse,PLZ,Ort,Telefon,Telefon2,Fax,Internet,Kdnr,Bemerkung,Gruppe,Land)
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

ButtonBeenden:
GuiEscape:
GuiClose:
ExitApp

ButtonAnrufen:
ControlGet, Zeile, List, Count Focused, SysListView321, Adressen
Loop % LV_GetCount("Col")
  LV_GetText(Row%a_index%,Zeile,a_index)
telefon:=Row6

2ButtonAnrufen:
Gui,2:Submit,NoHide
nr:=(a_guicontrol="tel1" or a_guicontrol="&Anrufen") ? telefon : (a_guicontrol="tel2") ? telefon2 : fax
nr:=regexreplace(nr,"[^\d#*]+")
if nr>
  run,ahk.exe C:\Eigene\prg\script\vcwdialf.ahk %nr%
return

2ButtonEtikett:
2ButtonEtikettBüwa:
Gui,2:Submit,NoHide
filedelete,%temp%\adr_drk.prn
ypos:=160
ypos:=(name2>"") ? ypos+40 : ypos
ypos:=(a_guicontrol="&Etikett BüWa") ? ypos+40 : ypos
out:="! 0 100 " . ypos . " 1`n"
if (privat=1)
  out.="TEXT 1 250 0 Absender privat...`n"
else
  out.="TEXT 1 250 0 Absender Geschäft...`n"
ypos:=10
if (a_guicontrol="&Etikett BüWa") {
  ypos+=40
  out.="TEXT 2 250 " . ypos . " BÜWA`n"
}
ypos+=40
out.="TEXT 2 250 " . ypos . " " . name . "`n"
if (name2>"") {
  ypos+=40
  out.="TEXT 2 250 " . ypos . " " . name2 . "`n"
}
ypos+=40
out.="TEXT 2 250 " . ypos . " " . strasse . "`n"
ypos+=40
out.="TEXT 2 250 " . ypos . " " . plz . " " . ort . "`n"
out.="END`n"
out:=ansi2ascii(out)
fileappend,%out%,%temp%\adr_drk.prn
msgbox ToDo: direkte Druckeransteuerung`n`n%temp%\adr_drk.prn erstellt mit Daten für Cognitive CL-422:`n%out%
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

ende:
if errorlevel {
  ListLines
  msgbox Fehler:%errorlevel%
}
exitapp

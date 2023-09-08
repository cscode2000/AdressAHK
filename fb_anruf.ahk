fritzbox=http://192.168.2.1:49000  ; Voraussetzung: Heimnetz -> Netzwerk -> Heimnetzfreigaben -> Zugriff für Anwendungen zulassen
user=User
kenn=Kennwort

Menu,Tray,Icon,c:\windows\system32\shell32.dll,197
AutoTrim, Off
nummer=%1%%2%%3%%4%%5%
if instr(nummer,"silent") or instr(nummer,"-s")
  silent=1
nummer:=regexreplace(nummer,"[^\d#*]+") ; nur Ziffern und *# erlaubt
tmpfile=%temp%\fritz_tmp.xml
location=/upnp/control/x_voip
urn1=dslforum-org
urn2=X_VoIP
uri=urn:%urn1%:service:%urn2%:1
action=X_AVM-DE_DialNumber
content=<NewX_AVM-DE_PhoneNumber>%nummer%</NewX_AVM-DE_PhoneNumber>
xmlstring=DialNumberResponse

soap=
(
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/" xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body>
<u:%action% xmlns:u="urn:%urn1%:service:%urn2%:1">%content%</u:%action%>
</s:Body></s:Envelope>
)

filedelete,%tmpfile%
fileappend,%soap%,%tmpfile%
runwait,curl.exe -k -m 5 --anyauth -o %tmpfile% -u "%user%:%kenn%" %fritzbox%%location% -H "Content-Type: text/xml`; charset=""utf-8""" -H "SoapAction: ""%uri%#%action%""" -d @%tmpfile%,,hide useerrorlevel
if silent>
  exit
fileread,f1,%tmpfile%
regexmatch(f1,xmlstring . ".+?" . xmlstring,soap)
meldung=Nummer wird gewählt: %nummer%`n`nBitte Telefon abheben!
if (soap<>"DialNumberResponse xmlns:u=""urn:dslforum-org:service:X_VoIP:1""></u:X_AVM-DE_DialNumberResponse")
  meldung.="`n`nResult:`n" soap
if instr(f1,xmlstring)
  msgbox,64,Fritzbox Dialer,%meldung%,10
else
{
  regexmatch(f1,"is)(<errorcode.+Description>)",soap)
  msgbox,48,Fritzbox Dialer,Fehler beim Wählen der Nummer: %nummer%`n`n%soap%
}

DIM objFS, objShell
SET objShell = CreateObject("WScript.Shell")
DIM strName, strKey, strPId, strOSType, strPath
Dim objFSO, objFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const Create = true
Dim cAuthor, iIHnr
strPath = objShell.ExpandEnvironmentStrings( "%SystemDrive%\TEMP\mail.TRP" )

''Arguments Aufruf und FormatierungMm
''Argument1 = IH (XXXXX) ; Argument2 = ggf. Authorenk�rzel
Dim strTRP
Set strTRP = WScript.arguments
if len(strTRP(0)<6) then
iIHnr = String(6-len(strTRP(0)),"0") & strTRP(0)	''IH Nummer muss Sechsstellig sein, Erste Ziffer unbedingt eine 0
end if
cAuthor = ucase(Left(strTRP(1),1))	''Authorenk�rzel darf nur ein Zeichen lang und Gro� sein

''Haupt Programm
Save("IA#7")	''Versions Hinweis f�r das mit den Daten weiterarbeitende Programm
strName = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName") ''Liest das Betriebsystem aus
strOSType = objShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE") ''Liest die installierte Architektur vom Betriebssystem aus
IF Len(strName) > 0 THEN
   strKey = DecodeKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId") ''Liest die ProductID aus zum dekodieren des Lizenzschl�ssels
	Dim strTempOS
	If instr(strName,"XP") > 0 then
			strTempOS = "XP"
		elseif instr(strName,"7") > 0 then
			strTempOS = "7"
		elseif instr(strName,"Vista") > 0 then
			strTempOS = "Vista"
		elseif instr(strName,"8") > 0 then
			strTempOS = "8"
	End If
    Select Case strOSType
		Case "x86" Save(iIHnr & " " & date & " ~   " & cAuthor & " OS=Win" & strTempOS & "_32")
		Case "AMD64" Save(iIHnr & " " & date & " ~   " & cAuthor & " OS=Win" & strTempOS & "_64")
	End Select
   Save(iIHnr & " " & date & " ~   " & cAuthor & " productkey=" & strKey)
   GetNWInfos
END IF
''Ende des Programms

'''FUNKTIONEN
'' Inhalt in Datei anh�ngen, wenn Datei nicht vorhanden, wird diese erstellt
Function Save(Content)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strPath, ForAppending, True)
	objFile.WriteLine Content
	objFile.close
	Set objFile = nothing
	Set objFSO = nothing
END Function

''ProduktId decodieren (NICHT VER�NDERN)
Function DecodeKey(RegKey)
   BinKey = objShell.RegRead(RegKey)
   CONST KeyOffset = 52
   iLen = 28
   szChars = "BCDFGHJKMPQRTVWXY2346789"
   DO
     x = 0
     n = 14
     DO
       x = x * 256
       x = BinKey(n + KeyOffset) + x
       BinKey(n + KeyOffset) = (x \ 24) and 255
       x = x Mod 24
       n = n - 1
     LOOP WHILE n >= 0
     iLen = iLen - 1
     szProductKey = mid(szChars, x + 1 ,1 ) & szProductKey
     if (((29 - iLen) Mod 6) = 0) and (iLen <> -1) then
       iLen = iLen - 1
       szProductKey = "-" & szProductKey
     END IF
   LOOP WHILE iLen >= 0
   DecodeKey = szProductKey
END Function

''Mac Adresse und zugeh�riges Interface
Function GetNWInfos
	On Error Resume Next 
	strComputer = "." 
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
	Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter") 
	For Each objItem in colItems
	IF Not objItem.MACAddress = "" and objItem.AdapterTypeID = 0 then
		Save(iIHnr & " " & date & " ~   " & cAuthor & " MAC=" & objItem.MACAddress)
	END IF
	Next
END Function 

DIM objFS, objShell
SET objShell = CreateObject("WScript.Shell")
DIM strName, strKey, strPId, strOSType, strPath
Dim objFSO, objFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const Create = true

strPath = objShell.ExpandEnvironmentStrings( "%SystemDrive%\TEMP\mail.trp" )
''Dim strFormatVersion as String = "IA#7"
dim intTVar

''Arguments Aufruf
''Argument0 = IH (XXXXX) ; Argument1 = ggf. Authorenkürzel
Dim strTRP
Set strTRP = WScript.arguments

'' Prüft nach Hilfe oder undefinierte Argumente
if instr(strTRP(0),"") > 0 or instr(strTRP(1),"") > 0 then
  strName = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
  strKey = DecodeKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")
  MsgBox"Produkt: " & vbTab & vbTab & strName & vbCrLf & "Produktschlüssel: " & vbTab & strKey, vbOKOnly or vbInformation, "Produktname und -schlüssel"
  WScript.Quit
elseif instr(strTRP(0),"help") > 0 or instr(strTRP(0),"?") > 0  then
  MsgBox "Wird das Programm ohne Parameter aufgerufen," & vbCrLf & "erscheint der Lizencode des Windows Systems"  & vbCrLf & "ansonsten ist die Syntax: " & vbCrLf & "CaptureSys.vbs [ID-Nr] [Autorekuerzel]" & vbCrLf & "CaptureSys.vbs /help -> Diese Hilfe" ,0,"CaptureSys - Hilfe"
  WScript.Quit
end if

strPath = objShell.ExpandEnvironmentStrings( "%SystemDrive%\TEMP\" & strTRP(0) & ".trp" )

''Haupt Programm
strName = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
strOSType = objShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
IF Len(strName) > 0 THEN
   strKey = DecodeKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")
   ''System("echo IA#7 > " & strPath)
   ''Save(strFormatVersion)
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
		Case "x86" Save(strTRP(0) & " " & date & " ~   " & strTRP(1) & " OS=Win" & strTempOS & "_32")
		Case "AMD64" Save(strTRP(0) & " " & date & " ~   " & strTRP(1) & " OS=Win" & strTempOS & "_64")
	End Select
   Save(strTRP(0) & " " & date & " ~   " & strTRP(1) & " productkey=" & strKey)
   GetNWInfos
END IF
''Ende des Programms

'''FUNKTIONEN
'' Inhalt in Datei anhängen, wenn Datei nicht vorhanden, wird diese erstellt
Function Save(Content)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strPath, ForAppending, True)
	objFile.WriteLine Content
	objFile.close
	Set objFile = nothing
	Set objFSO = nothing
END Function

''ProduktId decodieren
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

''Mac Adresse und zugehöriges Interface
Function GetNWInfos
	On Error Resume Next 
	strComputer = "." 
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
	Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter") 
	For Each objItem in colItems
	IF Not objItem.MACAddress = "" and objItem.AdapterTypeID = 0 then
	''IF objItem.AdapterTypeID = 1 then
		''Save(strTRP(0) & " " & date & " I   " & strTRP(1) & " " & objItem.Description)
		if intTVar < 1 then
		Save(strTRP(0) & " " & date & " ~   " & strTRP(1) & " MAC=" & objItem.MACAddress)
		elseif intTVar =>1 then
		Save(strTRP(0) & " " & date & " ~   " & strTRP(1) & " MAC" & intTVar & "=" & objItem.MACAddress)
		end if
		intTVar=intTVar + 1
	''END IF 
	END IF
	Next
END Function 

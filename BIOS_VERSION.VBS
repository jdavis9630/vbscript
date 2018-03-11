'VBSCRIPT
'CHECK LIST OF COMPUTERS FOR BIOS VERSION
'RUN AS ADMINISTRATOR
'FOLLOW INSTRUCTIONS IN COMMENTS BELOW

On Error Resume Next 
Const ForReading = 1 
Set objFSO = CreateObject("Scripting.FileSystemObject") 

MsgBox "Script Executing. Please wait." 

'ENTER THE PATH TO THE FILE THAT CONTAINS YOUR LIST OF COMPUTERS
Set objTextFile = objFSO.OpenTextFile("C:\<PATH>\comp.txt", 1)

'ENTER THE PATH WHERE YOU WANT TO SAVE YOUR OUTPUT
Set outfile = objFSO.CreateTextFile("C:\<PATH>\BIOS_VERSION.txt") 

Do Until objTextFile.AtEndOfStream  
    strComputer = objTextFile.Readline 
	OutFile.WriteLine "========================================================"
	OutFile.WriteLine strComputer
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	If Err.Number <> 0 Then
		OutFile.WriteLine "Error: " & Err.Number
		OutFile.WriteLine "Error (Hex): " & Hex(Err.Number)
		OutFile.WriteLine "Source: " &  Err.Source
		OutFile.WriteLine "Description: " &  Err.Description
		Err.Clear
	Else
		Set colItems = objWMIService.ExecQuery( _
			"SELECT * FROM Win32_ComputerSystem",,48) 
		For Each objItem in colItems 
			If isNull(objItem.Name) Then
				OutFile.WriteLine "Hostname: UNKNOWN"
			Else
				OutFile.WriteLine "Hostname: " & objItem.Name
			End If
		Next

		If Err.Number <> 0 Then
			OutFile.WriteLine "Error: " & Err.Number
			OutFile.WriteLine "Error (Hex): " & Hex(Err.Number)
			OutFile.WriteLine "Source: " &  Err.Source
			OutFile.WriteLine "Description: " &  Err.Description
			Err.Clear
		End If
		Set colItems = objWMIService.ExecQuery( _
			"SELECT * FROM Win32_BIOS",,48) 
		For Each objItem in colItems 
		If isNull(objItem.SMBIOSBIOSVersion) Then
				OutFile.WriteLine "SMBIOSBIOSVersion: UNKNOWN"
			Else
				OutFile.WriteLine "SMBIOSBIOSVersion: " & objItem.SMBIOSBIOSVersion
			End If

		Next
	End If
Loop
objTextFile.Close
msgbox("Script Finished")
Debug.WriteLine("")
WScript.Quit

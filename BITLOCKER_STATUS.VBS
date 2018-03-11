'VBSCRIPT
'FIND BITLOCKER STATUS
'RUN AS ADMIN
'FOLLOW INSTRUCTIONS IN COMMENTS BELOW

On Error Resume Next 
Const ForReading = 1 
Set objFSO = CreateObject("Scripting.FileSystemObject") 

MsgBox "Click OK to begin running script" 

'ENTER THE PATH TO THE FILE THAT CONTAINS YOUR LIST OF COMPUTERS
Set objTextFile = objFSO.OpenTextFile("C:\<PATH>\comp.txt", 1)

'ENTER THE PATH WHERE YOU WANT TO SAVE YOUR OUTPUT
Set outfile = objFSO.CreateTextFile("C:\<PATH>\BITLOCKER_STATUS.txt") 

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
			"SELECT * FROM Win32_EncryptableVolume",,48) 
		For Each objItem in colItems 
			If isNull(objItem.DeviceID) Then
				OutFile.WriteLine "DeviceID: UNKNOWN"
			Else
				OutFile.WriteLine "DeviceID: " & objItem.DeviceID
			End If
			If isNull(objItem.DriveLetter) Then
				OutFile.WriteLine "DriveLetter: UNKNOWN"
			Else
				OutFile.WriteLine "DriveLetter: " & objItem.DriveLetter
			End If
			If isNull(objItem.PersistentVolumeID) Then
				OutFile.WriteLine "PersistentVolumeID: UNKNOWN"
			Else
			   OutFile.WriteLine "PersistentVolumeID: " & objItem.PersistentVolumeID
			End If
			If isNull(objItem.ProtectionStatus) Then
				OutFile.WriteLine "ProtectionStatus: UNKNOWN"
			Else
				OutFile.WriteLine "ProtectionStatus: " & objItem.ProtectionStatus
			End If
		Next
	End If
Loop
objTextFile.Close
msgbox("Script Finished")
Debug.WriteLine("")
WScript.Quit

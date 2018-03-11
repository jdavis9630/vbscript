'VBSCRIPT
'FIND BOOT CONFIGURATION FROM LIST OF COMPUTERS
'RUN AS ADMINISTRATOR
'FOLLOW INSTRUCTIONS IN COMMENTS BELOW

On Error Resume Next 
Const ForReading = 1 
Set objFSO = CreateObject("Scripting.FileSystemObject") 

MsgBox "Click OK to begin running script." 

'ENTER THE PATH TO THE FILE THAT CONTAINS YOUR LIST OF COMPUTERS
Set objTextFile = objFSO.OpenTextFile("C:\<PATH>\comp.txt", 1)

'ENTER THE PATH WHERE YOU WANT TO SAVE YOUR OUTPUT
Set outfile = objFSO.CreateTextFile("C:\<PATH>\BOOT_CONFIG.txt") 

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
			"SELECT * FROM Win32_BootConfiguration",,48) 
		For Each objItem in colItems 
			If isNull(objItem.BootDirectory) Then
				OutFile.WriteLine "BootDirectory: UNKNOWN"
			Else
				OutFile.WriteLine "BootDirectory: " & objItem.BootDirectory
			End If

			If isNull(objItem.Caption) Then
				OutFile.WriteLine "Caption: UNKNOWN"
			Else
				OutFile.WriteLine "Caption: " & objItem.Caption
			End If

			If isNull(objItem.ConfigurationPath) Then
				OutFile.WriteLine "ConfigurationPath: UNKNOWN"
			Else
				OutFile.WriteLine "ConfigurationPath: " & objItem.ConfigurationPath
			End If
			If isNull(objItem.Description) Then
				OutFile.WriteLine "Description: UNKNOWN"
			Else
				OutFile.WriteLine "Description: " & objItem.Description
			End If

			If isNull(objItem.LastDrive) Then
				OutFile.WriteLine "LastDrive: UNKNOWN"
			Else
				OutFile.WriteLine "LastDrive: " & objItem.LastDrive
			End If

			If isNull(objItem.Name) Then
				OutFile.WriteLine "Name: UNKNOWN"
			Else
				OutFile.WriteLine "Name: " & objItem.Name
			End If

			If isNull(objItem.ScratchDirectory) Then
				OutFile.WriteLine "ScratchDirectory: UNKNOWN"
			Else
				OutFile.WriteLine "ScratchDirectory: " & objItem.ScratchDirectory
			End If

			If isNull(objItem.SettingID) Then
				OutFile.WriteLine "SettingID: UNKNOWN"
			Else
				OutFile.WriteLine "SettingID: " & objItem.SettingID
			End If
			If isNull(objItem.TempDirectory) Then
				OutFile.WriteLine "TempDirectory: UNKNOWN"
			Else
				OutFile.WriteLine "TempDirectory: " & objItem.TempDirectory
			End If
		Next
	End If
Loop
objTextFile.Close
msgbox("Script Finished")
Debug.WriteLine("")
WScript.Quit
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
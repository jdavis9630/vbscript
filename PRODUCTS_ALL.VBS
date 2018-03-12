'VBSCRIPT
'GET PRODUCT LIST FROM LIST OF COMPUTERS
'RUN AS ADMINISTRATOR
'FOLLOW INSTRUCTIONS IN COMMENTS BELOW

On Error Resume Next 
Const ForReading = 1 
Set objFSO = CreateObject("Scripting.FileSystemObject") 

MsgBox "Click OK to begin running script." 

'ENTER THE PATH TO THE FILE THAT CONTAINS YOUR LIST OF COMPUTERS
Set objTextFile = objFSO.OpenTextFile("C:\<PATH>\comp.txt", 1)

'ENTER THE PATH WHERE YOU WANT TO SAVE YOUR OUTPUT
Set outfile = objFSO.CreateTextFile("C:\<PATH>\PRODUCTS_ALL.txt") 

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
				OutFile.WriteLine "================"
			End If
		Next	
		Set colItems = objWMIService.ExecQuery( _
			"SELECT * FROM Win32_Product",,48) 
		For Each objItem in colItems 
			If isNull(objItem.Description) Then
				OutFile.WriteLine "Description: UNKNOWN"
			Else
				OutFile.WriteLine "Description: " & objItem.Description
			End If	
			If isNull(objItem.HelpLink) Then
				OutFile.WriteLine "HelpLink: UNKNOWN"
			Else
				OutFile.WriteLine "HelpLink: " & objItem.HelpLink
			End If	
			If isNull(objItem.IdentifyingNumber) Then
				OutFile.WriteLine "IdentifyingNumber: UNKNOWN"
			Else
				OutFile.WriteLine "IdentifyingNumber: " & objItem.IdentifyingNumber
			End If					
			If isNull(objItem.InstallDate) Then
				OutFile.WriteLine "InstallDate: UNKNOWN"
			Else
				OutFile.WriteLine "InstallDate: " & objItem.InstallDate
			End If					
			If isNull(objItem.InstallSource) Then
				OutFile.WriteLine "InstallSource: UNKNOWN"
			Else
				OutFile.WriteLine "InstallSource: " & objItem.InstallSource
			End If					
			If isNull(objItem.LocalPackage) Then
				OutFile.WriteLine "LocalPackage: UNKNOWN"
			Else	
				OutFile.WriteLine "LocalPackage: " & objItem.LocalPackage
			End If					
			If isNull(objItem.PackageCache) Then
				OutFile.WriteLine "PackageCache: UNKNOWN"
			Else	
				OutFile.WriteLine "PackageCache: " & objItem.PackageCache
			End If					
			If isNull(objItem.PackageCode) Then
				OutFile.WriteLine "objItem.PackageCode: UNKNOWN"
			Else
				OutFile.WriteLine "PackageCode: " & objItem.PackageCode
			End If	
			If isNull(objItem.PackageName) Then
				OutFile.WriteLine "objItem.PackageName: UNKNOWN"
			Else			
				OutFile.WriteLine "PackageName: " & objItem.PackageName
			End If					
			If isNull(objItem.Vendor) Then
				OutFile.WriteLine "objItem.Vendor: UNKNOWN"
			Else
				OutFile.WriteLine "Vendor: " & objItem.Vendor
			End If					
			If isNull(objItem.Version) Then
				OutFile.WriteLine "objItem.Version: UNKNOWN"
			Else
			OutFile.WriteLine "Version: " & objItem.Version
			End If	
			OutFile.WriteLine "---------------------"
		Next
	End if
Loop
objTextFile.Close
msgbox("Script Finished")
Debug.WriteLine("")
WScript.Quit

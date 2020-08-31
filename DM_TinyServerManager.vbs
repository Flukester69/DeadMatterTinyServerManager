' ***********************************************************************************
' Title:   Dead Matter Server Manager written in lovely vb :)
' Author:  Flukey
' Discord: https://discord.gg/F2czUMD
' Version: 1.0
'
' Usage:   One line to modify in Main Dead Matter Section below. One line to modify.
' ***********************************************************************************

Function IsTheServerRunning( strProcess )
    Dim Process, strObject
	
    IsTheServerRunning = False
    strObject   = "winmgmts://" & "."
    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
	If UCase( Process.name ) = UCase( strProcess ) Then
            IsTheServerRunning = True
            Exit Function
        End If
    Next
	
End Function

Function KillTheServer( strProcess )

	Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcess &"'")

	For Each objProcess in colProcessList
		objProcess.Terminate()
	Next
	
End Function

Function CheckMemoryFree()

	Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)
	GB = 1024 *1024 * 1024
	For Each objItem in colItems
		FreeMEM = Round(objItem.AvailableBytes / GB,2)
		FreeMEM = FreeMEM & "GB"
	Next

End Function

Function StartTheServer()

	WshShell.Run strFullPath
		
End Function


' ****************************************************************************
' Main Dead Matter Stuff
' ****************************************************************************
'
Dim strComputer, strProcess, strServerPath, strFullPath, WshShell
Dim startupMsg, r_count, UserSaidNo

' YOU ONLY NEED TO MODIFY THE LINES BELOW
strServerPath = "C:\Steam\steamapps\common\Dead Matter Dedicated Server\deadmatterServer.exe"

' SERVER CHECK INTERVAL: Default is 60 seconds
ServerCheckInt = 60

'
' DON'T MODIFY ANYTHING BELOW.
'
strFullPath = chr(34) & strServerPath & chr(34) & " -log"
' current executable to monitor
strProcess = "deadmatterServer-Win64-Shipping.exe"

Set WshShell = CreateObject("WScript.Shell")
UserSaidNo = 1
r_count = 0
CrashTime = FormatDateTime(Now)

Do While startup <> vbCancel

	' Check Status of Server and restart if applicable.
	
	CheckMemoryFree

	If( IsTheServerRunning( strProcess ) = True ) Then

		If r_count = 0 Then
			CrashTime = "Never"
			startupMsg = "        Server Status: RUNNING " & vbcrlf & "              Free RAM: " & FreeMEM & vbcrlf & "                 Restarts: " & r_count & vbcrlf & "                       Time: " & CrashTime & vbcrlf & vbcrlf & "[Yes]=START - [No]= STOP - [Cancel]= Exit "
		Else
			startupMsg = "        Server Status: RUNNING " & vbcrlf & "              Free RAM: " & FreeMEM & vbcrlf & "                 Restarts: " & r_count & vbcrlf & "                       Time: " & CrashTime & vbcrlf & vbcrlf & "[Yes]=START - [No]= STOP - [Cancel]= Exit "
		End If

	Else

		If UserSaidNo = 0 Then
			r_count = r_count + 1
			CrashTime = FormatDateTime(Now)
			startupMsg = "        Server Status: STOPPED " & vbcrlf & "              Free RAM: " & FreeMEM & vbcrlf & "                 Restarts: " & r_count & vbcrlf & "                       Time: " & CrashTime & vbcrlf & vbcrlf & "[Yes]=START - [No]= STOP - [Cancel]= Exit "
			StartTheServer
		Else
			If r_count = 0 Then
				CrashTime = "Never"
			Else
				CrashTime = FormatDateTime(Now)
			End If
			
			If FreeMem = "" and UserSaidNo = 1 Then
				FreeMem = "n/a"
			End If
			
			startupMsg = "        Server Status: STOPPED " & vbcrlf & "              Free RAM: " & FreeMEM & vbcrlf & "                 Restarts: " & r_count & vbcrlf & "                       Time: " & CrashTime & vbcrlf & vbcrlf & "[Yes]=START - [No]= STOP - [Cancel]= Exit "
		End If
		
	End If

	' Show UI for Management
	' server-check-interval is every 60 seconds for now
	startup = WshShell.Popup(startupMsg, ServerCheckInt, "Dead Matter Tiny Server Manager", 3 + 32)

	Select Case startup

		Case 6
			' If user selects Yes. Then launch.
			If( IsTheServerRunning( strProcess ) = True ) Then
				UserSaidNo = 0
				MsgBox "Server is already RUNNING! "
			Else
				StartTheServer
				UserSaidNo = 0			
			End If
		
		Case 7
			' If user selects No. Then exit.
			UserSaidNo = 1
			
			If( IsTheServerRunning( strProcess ) = True ) Then
				ATTResponse = _
					Msgbox("Server is RUNNING are you sure? ", _
					vbYesNo, "Please Confirm! ")
					If ATTResponse = vbYes Then
						KillTheServer(strProcess)
						Wscript.Sleep(5000)
					End If
			Else
				MsgBox "Oops! Server isn't currently Running! "
			End If

		Case 2
			startup = vbCancel
			
		End Select
		
Loop
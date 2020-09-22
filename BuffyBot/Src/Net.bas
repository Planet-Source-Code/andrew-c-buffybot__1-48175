Attribute VB_Name = "Net"
'// Net RTF functions v1

Public Sub ConnectServer(ServerString As String)
Dim serverip As String
Dim serverport As String
If InStr(1, ServerString, ":", vbTextCompare) <> 0 Then
    serverip = Left(ServerString, InStr(1, ServerString, ":", vbTextCompare) - 1)
    serverport = Right(ServerString, Len(ServerString) - Len(serverip) - 1)
End If
If InStr(1, ServerString, " ", vbTextCompare) <> 0 Then
    serverip = Left(ServerString, InStr(1, ServerString, " ", vbTextCompare) - 1)
    serverport = Right(ServerString, Len(ServerString) - Len(serverip) - 1)
End If
If Val(serverport) > 65535 Then
    ' The port specified was invalid. must be greater than
    ' or equal to 65535 (max long)
    AppendConsole "* Invalid server port: " & serverport
    Exit Sub
End If
If serverip = "" Then serverip = ServerString: serverport = "6667" ' Just use some basic defaults
frmMain.IRCSocket.Close
frmMain.IRCSocket.Connect serverip, serverport
AppendConsole "*** Connecting to " & serverip & " on port " & serverport
End Sub

Public Sub SendIrc(sString As String)
'If frmMain.IRCSocket.State = 7 Then
'    frmMain.IRCSocket.SendData sString
'End If
MessageQue.SpoolAdd sString
End Sub


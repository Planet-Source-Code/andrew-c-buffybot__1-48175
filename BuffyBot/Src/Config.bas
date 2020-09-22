Attribute VB_Name = "Config"
'// Buffybot's configuration routines

Public CFGVars As New Dictionary
Public Servers() As String
Public ServerStrings As String

Public Sub GetNextServer()
For x = 0 To UBound(Servers)
If Servers(x) = ServerStrings Then
    If x = UBound(Servers) Then ServerStrings = Servers(0): Exit Sub
    ServerStrings = Servers(x + 1)
    Exit Sub
End If
Next x
End Sub

Public Sub LoadConfig()
Dim configcfg As String
Dim headerflag As String
Dim sTemp As String
Dim fHandle As Long
fHandle = FreeFile
Open App.Path & "\config.ini" For Input As #fHandle
Do Until EOF(fHandle)
Line Input #fHandle, sTemp
' Remove comments
If InStr(1, sTemp, "#", vbTextCompare) <> 0 Then
    sTemp = Left(sTemp, InStr(1, sTemp, "#", vbTextCompare) - 1)
End If
sTemp = Trim(sTemp)
If Left(sTemp, 1) = "[" And Right(sTemp, 1) = "]" Then
    ' A header
    headerflag = Mid(sTemp, 2, Len(sTemp) - 2)
End If
' Is the format a set var?
If InStr(1, sTemp, " = ", vbTextCompare) <> 0 Then
    Dim varname As String, varvalue As String
    varname = Left(sTemp, InStr(1, sTemp, " ", vbTextCompare) - 1)
    varvalue = Right(sTemp, Len(sTemp) - Len(varname) - 3)
    CFGVars(varname) = varvalue
End If
' Add server?
'Stop
If LCase(Left(sTemp, Len("addserver"))) = "addserver" Then
    Dim ServerString As String
    ServerString = Right(sTemp, Len(sTemp) - Len("addserver "))
    If ServerStrings = "" Then
        ServerStrings = ServerString
    Else
        ServerStrings = ServerStrings & vbCrLf & ServerString
    End If
End If
If LCase(Left(sTemp, Len("goservers"))) = "goservers" Then
    Servers = Split(ServerStrings, vbCrLf)
End If
Loop
End Sub



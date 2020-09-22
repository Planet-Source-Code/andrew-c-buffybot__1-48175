Attribute VB_Name = "Channels"
Public ChannelObjects As New Collection


Public Sub LoadChannelFile()
Dim filehandle As Long
Dim sTemp As String
Dim WorkingChannel As String
Dim WorkingParam As String
filehandle = FreeFile
Open App.Path & "\Channels.cfg" For Input As #filehandle
Do Until EOF(filehandle)
Line Input #filehandle, sTemp
If InStr(1, sTemp, "//") <> 0 Then
    sTemp = Left(sTemp, InStr(1, sTemp, "//", vbTextCompare) - 1)
End If
If Left(sTemp, 8) = "<#!Start" Then
    ' Get the channel name out of the tag, then create a new
    ' channel class and add it to the collection
    Dim cName As String
    cName = Mid(sTemp, 10, Len(sTemp) - Len("<#!Start ") - 1)
    WorkingChannel = cName
    AddChanClass cName
End If
If Left(sTemp, 6) = "<#!End" Then WorkingChannel = ""
Loop
End Sub


Public Sub AddChanClass(ChannelName As String)
Dim newclass As ChanClass
Set newclass = New ChanClass
newclass.ChanName = ChannelName
ChannelObjects.Add newclass, ChannelName
End Sub

Private Sub SetChanVar(ChannelName As String, VarName As String, VarValue As String)
'Stop
For x = 1 To ChannelObjects.Count
If LCase(ChannelObjects(x).ChanName) = LCase(ChannelName) Then
    ChannelObjects(x).Attributes(VarName) = VarValue
    Exit Sub
End If
Next x
End Sub

Public Sub JoinChannels()
'Stop
For x = 1 To ChannelObjects.Count
MessageQue.SpoolAdd "JOIN " & ChannelObjects(x).ChanName & vbCrLf
Next x
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BotCoreFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Public BotNick As String
Public Realname As String
Dim fso As New FileSystemObject

Public Function RootPath()
RootPath = App.Path
End Function

Public Sub WriteConsole(sString As String)
AppendConsole sString
End Sub

Public Sub Terminate()
End
End Sub

Public Function LeftString(What As String, sLen As Long)
LeftString = Left(What, sLen)
End Function

Public Function RightString(What As String, sLen As Long)
RightString = Right(What, sLen)
End Function

Public Function UserExists(Username As String)
If fso.FileExists(App.Path & "\User Files\" & Username & ".dat") = False Then
    UserExists = 0
Else
    UserExists = 1
End If
End Function

Public Sub WriteUserValue(Username As String, sName As String, sVal As String)
If fso.FolderExists(App.Path & "\user Files") = False Then
    MkDir App.Path & "\User Files"
End If
IniFunctions.WriteINIVAL "UserFile", sName, sVal, App.Path & "\User Files\" & Username & ".dat"
End Sub

Public Sub WaitFor(waitint As Long)
' no thread to do this yet :(
End Sub

Public Function GetUserValue(Username As String, sName As String)
'Stop
If fso.FolderExists(App.Path & "\user Files") = False Then
    MkDir App.Path & "\User Files"
End If
' Does the userfile exist?
If fso.FileExists(App.Path & "\User Files\" & Username & ".dat") = False Then GetUserValue = -1: Exit Function
GetUserValue = IniFunctions.GetINIVal("UserFile", sName, App.Path & "\User Files\" & Username & ".dat")
End Function

Public Function VerifyFlag(Username As String, Channel As String, Flag As String)
Dim flagbuffer As String
flagbuffer = IniFunctions.GetINIVal(Channel, "ChanFlags", App.Path & "\User Files\" & Username & ".dat", -1)
If InStr(1, flagbuffer, Flag, vbTextCompare) <> 0 Then
    ' The flag is activated!
    VerifyFlag = 1
Else
    VerifyFlag = 0
End If
End Function

Public Sub ActivateFlag(Username As String, Channel As String, Flag As String)
Dim flagbuffer As String
flagbuffer = IniFunctions.GetINIVal(Channel, "ChanFlags", App.Path & "\User Files\" & Username & ".dat", "")
' is the flag already set?
If InStr(1, flagbuffer, Flag, vbTextCompare) <> 0 Then
    ' The flag is already activated!
    Exit Sub
Else
    flagbuffer = flagbuffer & Flag
    ' Now write the new flags buffer to the user file
    IniFunctions.WriteINIVAL Channel, "ChanFlags", flagbuffer, App.Path & "\User Files\" & Username & ".dat"
End If
End Sub

Public Sub DeactivateFlag(Username As String, Channel As String, Flag As String)
Dim flagbuffer As String
flagbuffer = IniFunctions.GetINIVal(Channel, "ChanFlags", App.Path & "\User Files\" & Username & ".dat", -1)
' is the flag set?
If InStr(1, flagbuffer, Flag, vbTextCompare) <> 0 Then
    ' The flag is activated, so remove it from the buffer
    flagbuffer = StringReplace(flagbuffer, Flag, "")
    ' Now write the new flags buffer to the user file
    IniFunctions.WriteINIVAL Channel, "ChanFlags", flagbuffer, App.Path & "\User Files\" & Username & ".dat"
End If
End Sub

Public Function GetChanFlag(Channel As String, Flag As String)
Dim FileHandle As Long
Dim sTemp As String
Dim WorkingChannel As String
Dim WorkingParam As String
'Stop
FileHandle = FreeFile
Open App.Path & "\Channels.cfg" For Input As #FileHandle
Do Until EOF(FileHandle)
Line Input #FileHandle, sTemp
If InStr(1, sTemp, "//") <> 0 Then
    sTemp = Left(sTemp, InStr(1, sTemp, "//", vbTextCompare) - 1)
End If
If Left(sTemp, 8) = "<#!Start" Then
    ' Get the channel name out of the tag, then create a new
    ' channel class and add it to the collection
    Dim cName As String
    cName = Mid(sTemp, 10, Len(sTemp) - Len("<#!Start ") - 1)
    WorkingChannel = cName
End If
If Left(sTemp, 6) = "<#!End" Then WorkingChannel = ""
If LCase(WorkingChannel) = LCase(Channel) Then
If Left(sTemp, 1) = "?" Then
    ' It's a variable setup
    Dim dstemp As String
    dstemp = Right(sTemp, Len(sTemp) - 1)
    If LCase(Left(dstemp, InStr(1, dstemp, " = ", vbTextCompare) - 1)) = LCase(Flag) Then
        GetChanFlag = Right(dstemp, Len(dstemp) - InStr(1, dstemp, " = ", vbTextCompare) - 2)
        Exit Function
    End If
End If
End If
Loop
End Function

Public Sub AddBan(Channel As String, Hostmask As String)
Stop
End Sub

Public Function CheckBan(Channel As String, Hostmask As String)
Dim FileHandle As Long
Dim sTemp As String
Dim WorkingChannel As String
Dim WorkingParam As String
'Stop
FileHandle = FreeFile
Open App.Path & "\Channels.cfg" For Input As #FileHandle
Do Until EOF(FileHandle)
Line Input #FileHandle, sTemp
If InStr(1, sTemp, "//") <> 0 Then
    sTemp = Left(sTemp, InStr(1, sTemp, "//", vbTextCompare) - 1)
End If
If Left(sTemp, 8) = "<#!Start" Then
    ' Get the channel name out of the tag, then create a new
    ' channel class and add it to the collection
    Dim cName As String
    cName = Mid(sTemp, 10, Len(sTemp) - Len("<#!Start ") - 1)
    WorkingChannel = cName
End If
If Left(sTemp, 6) = "<#!End" Then WorkingChannel = ""
If LCase(WorkingChannel) = LCase(Channel) Then
If Left(sTemp, 7) = "!addban" Then
    ' It's a variable setup
    Dim btemp As String
    btemp = Right(sTemp, Len(sTemp) - Len("!addban "))
    If Left(btemp, 1) = "~" Then btemp = Right(btemp, Len(btemp) - 1)
    If LCase(btemp) = LCase(Hostmask) Then
        CheckBan = 1
        Exit Function
    End If
End If
End If
Loop
End Function

Public Function GetNowTime$()
GetNowTime = Now
End Function

Public Function CheckUser(sUsername As String)
'Stop
If fso.FileExists(App.Path & "\User Files\" & sUsername & ".dat") = True Then CheckUser = True
End Function

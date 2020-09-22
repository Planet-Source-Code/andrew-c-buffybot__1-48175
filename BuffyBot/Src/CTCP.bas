Attribute VB_Name = "CTCP"
'// CTCP Module

Public Sub DealCTCPSubsystem(what As String, whosaid As String)
Dim userident As String
Dim sCommand As String
Dim sParam As String
userident = Left(whosaid, InStr(1, whosaid, "!", vbTextCompare) - 1)
what = Mid(what, 2, Len(what) - 2)
If InStr(1, what, " ", vbTextCompare) <> 0 Then
    sCommand = Left(what, InStr(1, what, " ", vbTextCompare) - 1)
    sParam = Right(what, Len(what) - Len(sCommand) - 1)
    Select Case LCase(sCommand)
    Case "ping"
    MessageQue.SpoolAdd "NOTICE " & userident & " :" & Chr(1) & "PING " & sParam & Chr(1) & vbCrLf
    Exit Sub
    Case "version"
    MessageQue.SpoolAdd "NOTICE " & userident & " :" & Chr(1) & "VERSION " & "Crano's IRC Client" & Chr(1) & vbCrLf
    End Select
Else
    Select Case LCase(what)
    Case "ping"
    MessageQue.SpoolAdd "NOTICE " & userident & " :" & Chr(1) & "PING" & Chr(1) & vbCrLf
    Exit Sub
    Case "version"
    MessageQue.SpoolAdd "NOTICE " & userident & " :" & Chr(1) & "VERSION " & "ALMIRC Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(1) & vbCrLf
    Case "time"
    MessageQue.SpoolAdd "NOTICE " & userident & " :" & Chr(1) & "TIME " & "The time by my watch is: " & Format(Now, "ddd dd/mm hh:mm:ss AM/PM" & Chr(1) & vbCrLf)
    End Select
    
End If

End Sub

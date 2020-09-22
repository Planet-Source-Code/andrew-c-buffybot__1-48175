Attribute VB_Name = "Console"
'// BuffyBot Console Module

Public Sub AppendConsole(sString As String, Optional ClearFirst As Boolean)
If ClearFirst = False Then
    frmMain.ConsoleLog.Text = frmMain.ConsoleLog.Text & vbCrLf & sString
Else
    frmMain.ConsoleLog.Text = sString
End If
frmMain.ConsoleLog.SelStart = Len(frmMain.ConsoleLog.Text)
End Sub

Public Sub ProcessLocalCommand(sString As String)
If Left(sString, 1) = "/" Then 'Command?
    sString = Right(sString, Len(sString) - 1)
Else
    Exit Sub
End If
' Single command?
If InStr(1, sString, " ", vbTextCompare) = 0 Then
    Select Case LCase(sString)
    Case "clear"
        frmMain.ConsoleLog.Text = ""
        Exit Sub
    End Select
Else
    ' Multi part command
        Dim bitf As String
        bitf = Right(sString, Len(sString) - InStr(1, sString, " ", vbTextCompare))
    Select Case LCase(Left(sString, InStr(1, sString, " ", vbTextCompare) - 1))
    Case "server"
    ' ok we wanna connect to a different server
    ' has a port been specified with the server?
    If InStr(1, bitf, " ", vbTextCompare) = 0 Then
        ' just the server, assume port 6667
        ConnectServer bitf & ":6667"
    Else
        ConnectServer bitf
    End If
    
    ' We wanna add a flag to a user's data file:
    Case "addflag"
        ' The bitf variable must hold the following:
        ' -
        ' 1. The nickname
        ' 2. Channel. Blank if global
        ' 3. The additional flag ie: +u. if they use -u then
        '    treat the command as a 'delflag' instead
        '    if |u| the command is global:
        Dim flagparams() As String
        flagparams = Split(bitf, " ")
        If UBound(flagparams) = 0 Then
            AppendConsole "* Invalid addflag command. Usage:"
            AppendConsole "  /addflag [nickname] [channel] [flag]"
            AppendConsole "  If channel is blank, then the flag is global"
            Exit Sub
        End If
        Dim sflag As String
        ' Are we serving the given user?
        If BotCore.UserExists(flagparams(0)) = 0 Then
            AppendConsole "* No registered user by that nick"
            Exit Sub
        End If
        If UBound(flagparams) = 1 Then
            ' A Global command?
            If Left(flagparams(1), 1) = "+" Or Left(flagparams(1), 1) = "-" Or Left(flagparams(1), 1) = "|" Then
                ' A global command:
                If Left(flagparams(1), 1) = "+" Then
                    sflag = Right(flagparams(1), 1)
                    sflag = Left(sflag, 1)
                    BotCore.ActivateFlag flagparams(0), "Global", sflag
                    Exit Sub
                End If
                If Left(flagparams(1), 1) = "-" Then
                    sflag = Right(flagparams(1), 1)
                    sflag = Left(sflag, 1)
                    BotCore.DeactivateFlag flagparams(0), "Global", sflag
                    Exit Sub
                End If
                If Left(flagparams(1), 1) = "|" Then
                    sflag = Right(flagparams(1), 1)
                    sflag = Left(sflag, 1)
                    BotCore.ActivateFlag flagparams(0), "Global", sflag
                    Exit Sub
                End If
              End If
        End If
        If UBound(flagparams) = 2 Then
            ' Stop here too
                If Left(flagparams(2), 1) = "+" Then
                    sflag = Right(flagparams(2), 1)
                    sflag = Left(sflag, 1)
                    BotCore.ActivateFlag flagparams(0), flagparams(1), sflag
                    Exit Sub
                End If
                If Left(flagparams(2), 1) = "-" Then
                    sflag = Right(flagparams(2), 1)
                    sflag = Left(sflag, 1)
                    BotCore.DeactivateFlag flagparams(0), flagparams(1), sflag
                    Exit Sub
                End If
                If Left(flagparams(2), 1) = "|" Then
                    sflag = Right(flagparams(2), 1)
                    sflag = Left(sflag, 1)
                    BotCore.ActivateFlag flagparams(2), flagparams(1), sflag
                    Exit Sub
                End If
        End If
        
    Case "setinfo"
    Dim infoparams() As String
    infoparams = Split(bitf, " ")
'    Stop
    If UBound(infoparams) = 0 Then
        ' Hmm... Just assume that they wanna erase their infoline
        BotCore.WriteUserValue infoparams(0), "Info", ""
        AppendConsole "* User info for " & infoparams(0) & " erased!"
        Exit Sub
    End If
    If UBound(infoparams) >= 1 Then
        infoline = Right(bitf, Len(bitf) - InStr(1, bitf, infoparams(1), vbTextCompare) + 1)
'        Stop
        BotCore.WriteUserValue infoparams(0), "Info", CStr(infoline)
        AppendConsole "* User info for " & infoparams(0) & " successfully updated!"
        Exit Sub
    End If
    End Select
End If
End Sub

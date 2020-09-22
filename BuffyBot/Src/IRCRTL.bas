Attribute VB_Name = "IRCRTL"
'// IRC Module
'// Handles the data from the IRC socket

Public Sub ProcessIRCData(sString As String)
Buffer = Buffer & sString
If InStr(1, Buffer, vbCrLf) = 0 Then temp = "": Exit Sub
temp = Mid$(Buffer, InStrRev(Buffer, vbCrLf) + 2)
Buffer = Mid$(Buffer, 1, Len(Buffer) - Len(temp))
Lines = Split(Buffer, vbCrLf)
Buffer = temp
On Local Error Resume Next
For x = 0 To UBound(Lines) - 1
    words = Split(Lines(x), " ")
    Debug.Print Lines(x)
    If Lines(x) = "" Then GoTo nextthing
    Select Case LCase(words(0))
         Case "ping"
            SendIrc "PONG " & words(1) & vbCrLf
        Case Else
    End Select
    If UBound(words) > 1 Then
        Select Case LCase(words(1))
            Case "nick"
            whosaid = Mid(words(0), 2)
                Nick = Mid(whosaid, 1, InStr(1, whosaid, "!") - 1)
                If BotCore.BotNick = Nick Then
                        BotCore.BotNick = Mid(words(2), 2)
                Else
                        ' Should we try and regain the nickname?
                        If LCase(Nick) = LCase(CFGVars("nickname")) Then
                            If LCase(CFGVars("regainnick")) = "yes" Then
                                'BotCore.BotNick = CFGVars("nickname")
                                SendIrc "NICK " & CFGVars("nickname") & vbCrLf
                            End If
                        End If
                End If
            Case "quit"
                whosaid = Mid(words(0), 2)
                Nick = Mid(whosaid, 1, InStr(1, whosaid, "!") - 1)
                Ident = Mid(whosaid, InStr(1, whosaid, "!") + 1, InStr(1, whosaid, "@") - Len(Nick) - 2)
                Host = Mid(whosaid, InStr(1, whosaid, "@") + 1)
                If LCase(Nick) = BotCore.BotNick Then
                AppendConsole "* Quit from IRC network"
                End If
                'Scripting.Run "OnPart", nick, IDENT, host, where
          List.Remove whosaid
               
            Case "part"
                whosaid = Mid(words(0), 2)
                Nick = Mid(whosaid, 1, InStr(1, whosaid, "!") - 1)
                Ident = Mid(whosaid, InStr(1, whosaid, "!") + 1, InStr(1, whosaid, "@") - Len(Nick) - 2)
                Host = Mid(whosaid, InStr(1, whosaid, "@") + 1)
                Where = words(2)
                List(whosaid).Channels.Remove Where
                If List(whosaid).Channels.Count = 0 Then List.Remove whosaid
                If Nick = BotCore.BotNick Then
                AppendConsole "* Left room " & Where
                End If
                ScriptObject.RunAll "OnPart", whosaid, Nick, Ident, Host, Where
            Case "join"
                whosaid = Mid(words(0), 2)
                Nick = Mid(whosaid, 1, InStr(1, whosaid, "!") - 1)
                Ident = Mid(whosaid, InStr(1, whosaid, "!") + 1, InStr(1, whosaid, "@") - Len(Nick) - 2)
                Host = Mid(whosaid, InStr(1, whosaid, "@") + 1)
                Where = Mid(words(2), 2)
                If BotCore.BotNick = Nick Then
                    SendIrc "WHO " & Where & vbCrLf
                    AppendConsole "* Joined room " & Where
                End If
                List.Add whosaid, Nick, Ident, Host, Where, ""
                List.Item(whosaid).Channels.Add Where
                ScriptObject.RunAll "OnJoin", Nick, Ident, Host, Where
            Case "mode"
                If words(2) = BotCore.BotNick Then
                    'Connected
                    AppendConsole "* " & words(2) & " sets mode " & Right(words(3), Len(words(3)) - 1)
                    JoinChannels
                End If
                
            Case "privmsg"
            
                whosaid = Mid(words(0), 2)
                Nick = Mid(whosaid, 1, InStr(1, whosaid, "!") - 1)
                Ident = Mid(whosaid, InStr(1, whosaid, "!") + 1, InStr(1, whosaid, "@") - Len(Nick) - 2)
                Host = Mid(whosaid, InStr(1, whosaid, "@") + 1)
                Where = words(2)
                what = Mid(Lines(x), InStr(2, Lines(x), ":") + 1)
            
            
                List.Add Nick & "!" & Ident & "@" & Host, Nick, Ident, Host, Where, ""
               
            

                If Left(Where, 1) <> "#" Then
                    If Left(Where, 1) = "&" Then GoTo chan
                    Inchannel = False
                    'BotCore.BotNick = Where
                    Where = Nick
                Else
chan:
                    Inchannel = True
                End If
                List.Add whosaid, Nick, Ident, Host, Where, ""
                List.Item(whosaid).Channels.Add Where
                If Left(what, 1) = Chr(1) And Right(what, 1) = Chr(1) Then
                    DealCTCPSubsystem CStr(what), CStr(whosaid)
                    ScriptObject.RunAll "CTCP", whosaid, what, Where, Ident, Host, Inchannel
                Else
                    ScriptObject.RunAll "PrivMsg", whosaid, what, Where, Ident, Host, Inchannel
                End If
            Case "433"
                
                GetAnotherNickname
            Case "332"
                BotCore.BotNick = words(2)
            Case "366"
                BotCore.BotNick = words(2)
            Case "333"
                BotCore.BotNick = words(2)
            Case "352"
                ':irc.ins.net.uk 352 VBBot #VisualBasic awindeyr 209.224.98.114 irc.ins.net.uk Adriana H@ :0 Adriana Windeyer
                 BotCore.BotNick = words(2)
                 List.Add words(7) & "!" & words(4) & "@" & words(5), words(7), words(4), words(5), words(3), words(8)
                 List.Item(words(7) & "!" & words(4) & "@" & words(5)).Channels.Add Where
                
        End Select
    End If
nextthing:
LastWhere = Where
Form1.Caption = List.Count
Next x
Buffer = ""
End Sub

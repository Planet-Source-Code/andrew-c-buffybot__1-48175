VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRCFuncts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// Classable irc functions
'// Revision 1.0

Public Sub SendIRC(sString As String)
MessageQue.SpoolAdd sString
End Sub


Public Sub UserNotice(Who As String, What As String)
'Stop
MessageQue.SpoolAdd "NOTICE " & Who & " :" & What & vbCrLf
End Sub

Public Sub PostMsg(Where As String, What As String)
MessageQue.SpoolAdd "PRIVMSG " & Where & " :" & What & vbCrLf
End Sub

Public Sub HandleSpecial(sNick As String, sString As String)
On Error Resume Next
Dim sText As String
Dim sType As String
sType = Left(sString, InStr(1, sString, "!", vbTextCompare) - 1)
sText = Right(sString, Len(sString) - Len(sType) - 1)
sText = StringReplace(sText, "%nickname%", sNick)
sText = StringReplace(sText, "%randnickname%", sNick)
Select Case sType
Case "Text"
    BotMain.IRCFunctions.PostMsg sNick, sText
Case "Action"
    BotMain.IRCFunctions.UserNotice sNick, sText
End Select
End Sub

Public Sub SetMode(Channel As String, Nick As String, Mode As String)
MessageQue.SpoolAdd "MODE " & Channel & " " & Mode & " " & Nick & vbCrLf
End Sub

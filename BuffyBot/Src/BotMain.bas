Attribute VB_Name = "BotMain"
Public ScriptObject As New ScriptClass
Public BotCore As New BotCoreFunctions
Public IRCFunctions As New IRCFuncts
Public SQLObject As New SQL
Public MathFunct As New MathFunctions
Public DataObj As New DataLibrary

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub RawLog(SString As String)
frmraw.txtraw.Text = frmraw.txtraw.Text & vbCrLf & SString
frmraw.txtraw.SelStart = Len(frmraw.txtraw.Text)
End Sub

Sub Main()
' Are we being run from the main directory?
Dim fso As New FileSystemObject
If fso.FileExists(App.Path & "\config.ini") = False Then
    MsgBox "Please run BuffyBot.exe from the main directory containing config.ini", vbCritical, "Runtime Error"
    End
End If
' Show the main form
frmMain.Show
frmraw.Show
AppendConsole "BuffyBot IRC Interface", True
AppendConsole "BuffyEngine version " & App.Major & "." & App.Minor & "." & App.Revision
AppendConsole "--------------------------------"
AppendConsole "Starting Init:"
Init
End Sub

Private Sub Init()
'// Load in scriptable components
AppendConsole "  -> Loading scripts..."
ScriptObject.Include "scripts\main.c"
ScriptObject.ScriptComponents(1).ExecuteStatement ("Main")
'// Load the configuration file
Config.LoadConfig
'// Setup a few other things
BotCore.BotNick = Config.CFGVars("nickname")
BotCore.Realname = Config.CFGVars("realname")

'// Load channels
Channels.LoadChannelFile

'// Start the identD server
frmMain.IdentSock(0).Listen
'// Now try a server
Config.ServerStrings = Config.Servers(0)
ConnectServer Config.ServerStrings
End Sub

Public Sub GetAnotherNickname()
Dim anothername As String
anothername = CFGVars("altnick")
Do While InStr(1, anothername, "?", vbTextCompare) <> 0
anothername = Replace(anothername, "?", Int(Rnd * 9))
Loop
'CFGVars("nickname") = anothername
BotCore.BotNick = anothername
SendIrc "NICK " & BotCore.BotNick & vbCrLf
SendIrc "USER " & BotCore.BotNick & " Userner Server : " & BotCore.Realname & vbCrLf & "MODE " & BotCore.BotNick & " +i " & vbCrLf
End Sub

Public Function StringReplace(SString As String, SearchFor As String, ReplaceWith As String)
Do While InStr(1, SString, SearchFor, vbTextCompare) <> 0
SString = Replace(SString, SearchFor, ReplaceWith)
Loop
StringReplace = SString
End Function

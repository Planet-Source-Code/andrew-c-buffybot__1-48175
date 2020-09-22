VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BuffyBot Console"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4905
      Top             =   1845
   End
   Begin VB.Timer ReconnectTimer 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   2400
      Top             =   1350
   End
   Begin MSScriptControlCtl.ScriptControl ScriptLib 
      Index           =   0
      Left            =   90
      Top             =   1815
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin MSWinsockLib.Winsock IdentSock 
      Index           =   0
      Left            =   5490
      Top             =   1545
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   113
   End
   Begin MSWinsockLib.Winsock IRCSocket 
      Left            =   5490
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Settings"
      Height          =   300
      Left            =   3240
      TabIndex        =   3
      Top             =   2850
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      Height          =   300
      Left            =   4170
      TabIndex        =   2
      Top             =   2850
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Height          =   2925
      Left            =   -15
      TabIndex        =   1
      Top             =   -105
      Width           =   6015
      Begin VB.TextBox ConsoleInput 
         Height          =   285
         Left            =   45
         TabIndex        =   5
         Top             =   2580
         Width           =   5925
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00800000&
         Height          =   2445
         Left            =   45
         ScaleHeight     =   2385
         ScaleWidth      =   5880
         TabIndex        =   4
         Top             =   120
         Width           =   5940
         Begin VB.TextBox ConsoleLog 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   2310
            Left            =   30
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   30
            Width           =   5805
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   300
      Left            =   5100
      TabIndex        =   0
      Top             =   2850
      Width           =   885
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function GetIdentSock() As Long
On Error Resume Next
For x = 0 To IdentSock.Count - 1
If IdentSock(x).State <> 7 And IdentSock(x).State <> 2 Then
    IdentSock(x).Close
    GetIdentSock = x
End If
Next x
Dim f As Long
f = IdentSock.Count + 1
Load IdentSock(f)
GetIdentSock = f
End Function


Private Sub ConsoleInput_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Left(ConsoleInput.Text, 1) = "/" Then
        ProcessLocalCommand ConsoleInput.Text
        ConsoleInput.Text = ""
        KeyAscii = 0
        Exit Sub
    End If
    SendIrc ConsoleInput.Text & vbCrLf
    KeyAscii = 0
    ConsoleInput.Text = ""
End If
End Sub

Private Sub IdentSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
IdentSock(GetIdentSock).Accept requestID
End Sub

Private Sub IdentSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
IdentSock(Index).GetData strdata
IdentSock(Index).SendData strdata & ":USERID:UNIX:" & CFGVars("userid") & vbCrLf
IdentSock(Index).Close
End Sub

Private Sub IRCSocket_Close()
AppendConsole "*** Disconnected"
End Sub

Private Sub IRCSocket_Connect()
AppendConsole "*** Connected to server"
' Send some nickname info
DoEvents
SendIrc "NICK " & BotCore.BotNick & vbCrLf
SendIrc "USER " & BotCore.BotNick & " Userner Server : " & BotCore.Realname & vbCrLf & "MODE " & BotCore.BotNick & " +i " & vbCrLf
'SendIrc "USER " & Config.CFGVars("userid") & " " & """" & Config.CFGVars("email") & """" & " " & """" & frmMain.IRCSocket.LocalHostName & """" & " :" & CFGVars("fullname") & vbCrLf
End Sub

Private Sub IRCSocket_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
IRCSocket.GetData strdata
RawLog strdata
ProcessIRCData strdata
End Sub

Private Sub IRCSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
AppendConsole "*** Socket Error (" & Number & "): " & Description
ReconnectTimer.Enabled = True
End Sub

Private Sub ReconnectTimer_Timer()
ReconnectTimer.Enabled = False
Config.GetNextServer
ConnectServer Config.ServerStrings
ReconnectTimer.Interval = ReconnectTimer.Interval
End Sub

Private Sub Timer1_Timer()
MessageQue.SendNext
End Sub

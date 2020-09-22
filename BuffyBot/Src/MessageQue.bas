Attribute VB_Name = "MessageQue"
'// Buffybot's IRC message que.
Public OutgoingQue(50) As String 'A 50 buffy message spool
Public IngoingQue(50) As String
Public Sub SpoolAdd(sString As String)
'Stop
For x = 0 To 50
If OutgoingQue(x) = "##########" Or OutgoingQue(x) = "" Then
    ' Use this spot
    OutgoingQue(x) = sString
    Exit Sub
End If
Next x
End Sub

Public Sub SendNext()
'Stop
If OutgoingQue(0) = "" Or OutgoingQue(0) = "##########" Then Exit Sub
'Stop
frmMain.IRCSocket.SendData OutgoingQue(0)
For x = 1 To 50
OutgoingQue(x - 1) = OutgoingQue(x) 'Move everything up one spot
If OutgoingQue(x - 1) = "" Then OutgoingQue(x - 1) = "##########"
Next x
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// Class file for math functions in scripting

Public Function ConvertTime(StoredTime As String)
i = DateDiff("s", StoredTime, Now)

    'extract number of days
    days = i \ 86400
    i = i Mod 86400
    
    'extract number of hrs
    hours = i \ 3600
    i = i Mod 3600
    
    'extract number of mins
    minutes = i \ 60
    i = i Mod 60
    
    'extract number of secs
    seconds = i
    
    

If days = 0 Then
    If hours = 0 Then
        If minutes = 0 Then
            If seconds = 0 Then
                '
            Else
                ConvertTime = seconds & " Seconds"
                Exit Function
            End If
        Else
            ConvertTime = minutes & " Minutes, " & seconds & " Seconds"
            Exit Function
        End If
    Else
        ConvertTime = hours & " Hours, " & minutes & " Minutes, " & seconds & " Seconds"
        Exit Function
    End If
Else
    ConvertTime = days & " Days, " & hours & " Hours, " & minutes & " Minutes, " & seconds & " Seconds"
    Exit Function
End If

'Stop
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// Class module to handle data files

Public Function Scandata(sString As String, Filename As String)
Dim FileHandle As Long
Scandata = 0
Dim temp As String
Dim temp2 As String
FileHandle = FreeFile
Open App.Path & "\Data\" & Filename For Input As #FileHandle
Do Until EOF(FileHandle)
Input #FileHandle, temp
If temp <> "" And temp <> " " And temp <> vbCrLf Then
    If InStr(1, sString, temp, vbTextCompare) <> 0 Then
        Scandata = 1
        Exit Function
    End If
End If
Loop
Close #FileHandle
End Function

Public Function GetRandomSentence(sFilename As String)
Dim d As String, e() As String
d = TextStreaming.ReadFile(App.Path & "\data\" & sFilename)
e = Split(d, vbCrLf)
Randomize
'Stop
rndt = Int(Rnd * UBound(e) - 1) + 1
GetRandomSentence = e(rndt)
End Function

Public Function GetRankedMsg(Header As String, Filename As String)
'Stop
Dim CurrentHeader As String
Dim Lines() As String
Dim TempString As String
Dim intd As String
Dim FileHandle As Long
FileHandle = FreeFile
Open App.Path & "\Data\" & Filename For Input As #FileHandle
Do Until EOF(FileHandle)
Line Input #FileHandle, TempString
If Left(TempString, Len("<Begin")) = "<Begin" Then
    CurrentHeader = Right(TempString, Len(TempString) - Len("<Begin "))
    CurrentHeader = Left(CurrentHeader, Len(CurrentHeader) - 1)
    GoTo foundheader:
End If
If Left(TempString, Len("<End")) = "<End" Then CurrentHeader = ""
If LCase(CurrentHeader) = LCase(Header) Then
    intd = intd & vbCrLf & TempString
End If

foundheader:
Loop
Close #FileHandle
Lines = Split(intd, vbCrLf)
Randomize
'Stop
randint = Int(Rnd * (UBound(Lines))) + 1
GetRankedMsg = Lines(randint) ' & " (" & randint & "/" & UBound(Lines) & ")"
End Function

Public Sub WriteDataFile(Filename As String, Header As String, Text As String)
Dim fileheader As String
Dim TempString As String
Dim AnotherTemp As String
Dim fso As New FileSystemObject
If fso.FileExists(App.Path & "\Data\" & Filename) = True Then
    ' Append to the data file structure
    Dim f1 As Long
    Dim f2 As Long
    f1 = FreeFile
    Open App.Path & "\Data\" & Filename For Input As #f1
    f2 = FreeFile
    Open App.Path & "\Data\" & "~1" & Filename For Output As #f2
    Do Until EOF(f1)
    Line Input #f1, TempString
    If TempString = "<End " & Header & ">" Then
        Print #f2, Text
        Print #f2, "<End " & Header & ">"
    Else
        Print #f2, TempString
    End If
    Loop
    Close #f2
    Close #f1
    Kill App.Path & "\Data\" & Filename
    Name App.Path & "\Data\" & "~1" & Filename As App.Path & "\Data\" & Filename
    Exit Sub
Else
    ' Create a new file structure
    fileheader = FreeFile
    Open App.Path & "\data\" & Filename For Output As #fileheader
    Print #fileheader, "<Begin " & Header & ">"
    Print #fileheader, Text
    Print #fileheader, "<End " & Header & ">"
    Close #fileheader
End If
End Sub

Public Function GetAll(Filename As String)
Dim df As Long
Dim linetemp As String
df = FreeFile
Filename = App.Path & "\data\" & Filename
Open Filename For Input As #df
Do Until EOF(df)
Line Input #df, linetemp
If Left(linetemp, Len("<Begin")) = "<Begin" Then GoTo skipme
If Left(linetemp, Len("<Begin")) = "<End" Then GoTo skipme
If GetAll = "" Then
    GetAll = linetemp
Else
    GetAll = GetAll & vbCrLf & linetemp
End If
skipme:
Loop
Close #df
End Function

Public Function GetLineCount(Filename As String)
Dim df As Long
Dim linetemp As String
Dim counter As Long
df = FreeFile
Filename = App.Path & "\data\" & Filename
Open Filename For Input As #df
Do Until EOF(df)
Line Input #df, linetemp
If Left(linetemp, Len("<Begin")) = "<Begin" Then GoTo skipme
If Left(linetemp, Len("<Begin")) = "<End" Then GoTo skipme
counter = counter + 1
skipme:
Loop
Close #df
GetLineCount = counter
End Function

Public Function GetLine(Filename As String, Line As Variant)
Dim linenumber As Long
Dim currentline As Long
Dim lH As Long
linenumber = CLng(Line)
currentline = 0
lH = FreeFile
Open App.Path & "\Data\" & Filename For Input As #lH
Do Until EOF(lH)
Line Input #lH, templine
If Line = currentline Then
    GetLine = templine
    Exit Function
End If
currentline = currentline + 1
Loop
Close #lH
End Function

Public Function FileExists(Filename As String)
Dim fso As New FileSystemObject
If fso.FileExists(App.Path & "\Data\" & Filename) = True Then
    FileExists = 1
Else
    FileExists = 0
End If
End Function

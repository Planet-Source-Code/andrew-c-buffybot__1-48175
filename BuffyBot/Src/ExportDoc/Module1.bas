Attribute VB_Name = "Module1"
Dim MainDir As String
Dim exportsubs() As String
Dim exportfunctions() As String
Dim subcount As Long
Dim functioncount As Long
Public Sub ProcessVBP(Filename As String)
On Error Resume Next
MainDir = Left(Filename, InStrRev(Filename, "\", -1) - 1)
Dim fileh As Long
fileh = FreeFile
Dim sTemp As String
Open Filename For Input As #fileh
Do Until EOF(fileh)
Line Input #fileh, sTemp
If sTemp = "" Then GoTo nextloop
If Left(sTemp, InStr(1, sTemp, "=", vbTextCompare) - 1) = "Class" Then
    ' Process the class file
    subcount = 0
    functioncount = 0
    ProcessClass Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=", vbTextCompare))
End If
nextloop:
Loop
Close #fileh

' now read all the temp files and compile a html file
Form1.File1.Path = "C:\"
Form1.File1.Pattern = "*.dat"
' Make the page header first:
Dim htmllong As Long
htmllong = FreeFile
Open App.Path & "\Exported.html" For Output As #htmllong
Print #htmllong, "<B>Report for exported functions in: " & Filename & "</B>"
Print #htmllong, "<P>"
Print #htmllong, "<B>Exported Sub Procedures:</B>"
Form1.File1.Refresh
Form1.File1.Pattern = "~subtemp*.dat"
For x = 0 To Form1.File1.ListCount - 1
Dim fsubname As String
Dim fsubproc As String
Dim fhandle As Long
fhandle = FreeFile
Open Form1.File1.Path & "\" & Form1.File1.List(x) For Input As #fhandle
Line Input #fhandle, fsubname
Line Input #fhandle, fsubproc
Close #fhandle
Kill Form1.File1.Path & "\" & Form1.File1.List(x)
Print #htmllong, "<br>"
Print #htmllong, fsubname & "(" & fsubproc & ")"
Next x
Print #htmllong, "</P>"
Print #htmllong, "<P>"
Print #htmllong, "<B>Exported Functions:</B>"
Form1.File1.Pattern = "~functiontemp*.dat"
Form1.File1.Refresh
For x = 0 To Form1.File1.ListCount - 1
fhandle = FreeFile
Open Form1.File1.Path & "\" & Form1.File1.List(x) For Input As #fhandle
Line Input #fhandle, fsubname
Line Input #fhandle, fsubproc
Close #fhandle
Kill Form1.File1.Path & "\" & Form1.File1.List(x)
Print #htmllong, "<br>"
Print #htmllong, fsubname & "(" & fsubproc & ")"


Next x
Print #htmllong, "</P>"
Close #htmllong

End Sub

Public Sub ProcessClass(Filename As String)
Dim classname As String
Dim classfile As String
classname = Left(Filename, InStr(1, Filename, ";", vbTextCompare) - 1)
classfile = Right(Filename, Len(Filename) - Len(classname) - 2)
Dim filehandle As Long
Dim templine As String
filehandle = FreeFile
Open MainDir & "\" & classfile For Input As #filehandle
Do Until EOF(filehandle)
Line Input #filehandle, templine
If LCase(Left(templine, Len("Public Sub"))) = "public sub" Then
    ' a publically exportable sub function
    Dim filehandle1 As Long
    filehandle1 = FreeFile
    Randomize
    subname = Left(templine, InStr(1, templine, "(", vbTextCompare) - 1)
    subname = Right(subname, Len(subname) - InStr(1, subname, "Sub ") - 3)
    subparas = Mid(templine, InStr(1, templine, "(", vbTextCompare) + 1, InStr(1, templine, ")", vbTextCompare))
    subparas = Left(subparas, Len(subparas) - 1)
    Open "C:\~subtemp" & Int(Rnd * 5555555) & ".dat" For Append As #filehandle1
    Print #filehandle1, subname
    Print #filehandle1, subparas
    Close #filehandle1
End If

If LCase(Left(templine, Len("Public function"))) = "public function" Then
    ' a publically exportable sub function
    Dim filehandle2 As Long
    filehandle2 = FreeFile
    Randomize
    subname = Left(templine, InStr(1, templine, "(", vbTextCompare) - 1)
    subname = Right(subname, Len(subname) - InStr(1, subname, "Function ") - 8)
    subparas = Mid(templine, InStr(1, templine, "(", vbTextCompare) + 1, InStr(1, templine, ")", vbTextCompare))
    subparas = Left(subparas, Len(subparas) - 1)
'    Stop
    Open "C:\~functiontemp" & Int(Rnd * 5555555) & ".dat" For Append As #filehandle2
    Print #filehandle2, subname
    Print #filehandle2, subparas
    Close #filehandle2
End If
Loop
End Sub

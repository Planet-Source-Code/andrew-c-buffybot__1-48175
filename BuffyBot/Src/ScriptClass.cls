VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// This class handles runtime scripting bits
Public ScriptComponents As New Collection

Public Sub Include(Filename As String)
On Error Resume Next
Dim codestream As String
codestream = TextStreaming.ReadFile(Filename)
If codestream = "<#!>No File<#!>" Then
    Console.AppendConsole "  -> Include Error: " & Filename
    Console.AppendConsole "  -> Description: File not found."
    Exit Sub
End If
' Create a new script object and add it to the collection
Dim ObjectLong As Long
ObjectLong = frmMain.ScriptLib.UBound + 1
Load frmMain.ScriptLib(ObjectLong)
' Add the required public classes to the scripting object

frmMain.ScriptLib(ObjectLong).AddObject "BotCore", BotMain.BotCore, True
frmMain.ScriptLib(ObjectLong).AddObject "Scripting", BotMain.ScriptObject, True
frmMain.ScriptLib(ObjectLong).AddObject "IRCFunctions", BotMain.IRCFunctions, True
frmMain.ScriptLib(ObjectLong).AddObject "SQL", BotMain.SQLObject, True
frmMain.ScriptLib(ObjectLong).AddObject "MathFunctions", BotMain.MathFunct, True
frmMain.ScriptLib(ObjectLong).AddObject "DataLibrary", BotMain.DataObj, True

frmMain.ScriptLib(ObjectLong).AddCode codestream 'Add code to object

If frmMain.ScriptLib(ObjectLong).Error.Number = 0 Then
    Console.AppendConsole "  -> Included file " & Filename & " (" & Len(codestream) & " bytes)"
    ' Add to the collection
    ScriptComponents.Add frmMain.ScriptLib(ObjectLong), CStr(ObjectLong)
Else
    Console.AppendConsole "  -> Include Error: " & Filename & " (" & Len(codestream) & " bytes)"
    Console.AppendConsole "  -> Description: " & frmMain.ScriptLib(ObjectLong).Error.Description
    Console.AppendConsole "  -> Line: " & frmMain.ScriptLib(ObjectLong).Error.Line
    ' Operation failed. ie: the object is fucked, so remove it from
    ' the array.
    Unload frmMain.ScriptLib(ObjectLong)
End If
End Sub

Public Sub ExecuteAll(sString As String)
On Error Resume Next
For x = 1 To ScriptComponents.Count
ScriptComponents(x).ExecuteStatement sString
Next x
End Sub

Public Sub RunAll(sCommand As String, ParamArray VarParams() As Variant)
On Error GoTo compileerror
Dim params() As Variant
params = VarParams
For x = 1 To frmMain.ScriptLib.Count - 1
DoEvents
' Does the active control house the procedure given?
For y = 1 To frmMain.ScriptLib(x).Procedures.Count
If LCase(sCommand) = LCase(frmMain.ScriptLib(x).Procedures(y)) Then GoTo foundok
Next y
GoTo nextitem
foundok:
'Stop
Select Case UBound(VarParams)
Case 1
frmMain.ScriptLib(x).Run sCommand, params(0), params(1)
Case 2
frmMain.ScriptLib(x).Run sCommand, params(0), params(1), params(2)
Case 3
frmMain.ScriptLib(x).Run sCommand, params(0), params(1), params(2), params(3)
Case 4
frmMain.ScriptLib(x).Run sCommand, params(0), params(1), params(2), params(3), params(4)
Case 5
frmMain.ScriptLib(x).Run sCommand, params(0), params(1), params(2), params(3), params(4), params(5)
Case 6
frmMain.ScriptLib(x).Run sCommand, params(0), params(1), params(2), params(3), params(4), params(5), params(6)
Case 7
frmMain.ScriptLib(x).Run sCommand, params(0), params(1), params(2), params(3), params(4), params(5), params(6), params(7)
Case 8
frmMain.ScriptLib(x).Run sCommand, params(0), params(1), params(2), params(3), params(4), params(5), params(6), params(7), params(8)
Case 9
frmMain.ScriptLib(x).Run sCommand, params(0), params(1), params(2), params(3), params(4), params(5), params(6), params(7), params(8), params(9)
End Select
nextitem:
Next x
Exit Sub
compileerror:
MsgBox "Compile Error:" & vbCrLf & vbCrLf & frmMain.ScriptLib(x).Error.Text & ":" & vbCrLf & frmMain.ScriptLib(x).Error.Description & vbCrLf & "Line: " & frmMain.ScriptLib(x).Error.Line
End Sub

Public Function GetProcedureCount()
For x = 0 To frmMain.ScriptLib.Count - 1
GetProcedureCount = GetProcedureCount + frmMain.ScriptLib(x).Procedures.Count
Next x
End Function

Public Function GetModuleCount()
For x = 0 To frmMain.ScriptLib.Count - 1
GetModuleCount = GetModuleCount + frmMain.ScriptLib(x).Modules.Count
Next x
End Function


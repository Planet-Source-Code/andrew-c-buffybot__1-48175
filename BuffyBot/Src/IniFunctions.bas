Attribute VB_Name = "IniFunctions"
'//
'// Filename: FileIO.bas
'// Description: Functions for reading/writing INI files
'//
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Public Sub WriteINIVAL(Section As String, Key As String, Value As String, Filename As String)
' Write a value to an INI-File
WritePrivateProfileString Section, Key, Value, Filename
End Sub

Public Function GetINIVal(sSection As String, sKey As String, Filename As String, Optional sDefault As String)
' Get a value from an INI
On Error Resume Next
Dim SString As String
SString = String(100, "*") 'Max 100 Chars
lLength = Len(SString)
sLength = GetPrivateProfileString(sSection, sKey, sDefault, SString, lLength, Filename)
SString = Left(SString, sLength) 'Trim the value
GetINIVal = SString
End Function



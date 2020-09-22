Attribute VB_Name = "TextStreaming"
'// Text stream IO module

Dim fso As New FileSystemObject

Public Sub TextOut(Filename As String, sString As String)
' This sub-function will create a new text stream and write data to it
Dim tx1 As TextStream
If fso.FileExists(Filename) = False Then
    fso.CreateTextFile Filename
End If
Set tx1 = fso.OpenTextFile(Filename, ForAppending)
'tx1 = fso.OpenTextFile(Filename)
tx1.WriteLine sString
tx1.Close
End Sub

Public Function ReadFile(Filename As String) As String
Dim tx1 As TextStream
If fso.FileExists(Filename) = False Then
    ReadFile = "<#!>No File<#!>"
End If
Set tx1 = fso.OpenTextFile(Filename, ForReading)
ReadFile = tx1.ReadAll
tx1.Close
End Function

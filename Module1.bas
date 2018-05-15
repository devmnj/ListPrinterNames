Attribute VB_Name = "Module1"

Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

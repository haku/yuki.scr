Attribute VB_Name = "modINI"

Option Explicit
Option Compare Binary
Option Base 0

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetFromIni(Section As String, Key As String, Directory As String) As String
Dim strBuffer As String
strBuffer = String(750, Chr(0))
Key$ = LCase$(Key$)
GetFromIni$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToIni(Section As String, Key As String, KeyValue As String, Directory As String)
Call WritePrivateProfileString(Section$, Key$, KeyValue$, Directory$)
End Sub

Public Function GetFromIniEx(Section As String, Key As String, sDef As String, sFile As String)
Dim a
a = GetFromIni(Section, Key, sFile)
If a = "" Then
    GetFromIniEx = sDef
Else
    GetFromIniEx = a
End If
End Function

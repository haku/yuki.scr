Attribute VB_Name = "modGenFnc"

Option Explicit
Option Compare Binary
Option Base 0

Sub AddToLog(txtLog As TextBox, ByVal a As String, Optional bNewLine As Boolean = True)
If bNewLine = True Then
    a = IIf(Len(txtLog.Text) > 0, vbNewLine, "") & Format(Now, "hh:mm:ss") & vbTab & a
End If
txtLog.SelStart = Len(txtLog.Text)
txtLog.SelText = a
txtLog.SelStart = Len(txtLog.Text)
End Sub

Function ConvertSecToMin(ByVal s As Long) As String
On Error GoTo ConvertSecToMin_err

Dim a As Long

a = (s Mod 60) 'get the seconds
ConvertSecToMin = Trim$(Str$(((s - a) / 60))) & ":" & Right$("0" & Trim$(Str$(a)), 2)

Exit Function
ConvertSecToMin_err:
ConvertSecToMin = ""
End Function

Function TrncFilePath(f As String, n As Long) As String
Dim i As Long, j As Long, x As Long

x = 1
i = -1

For j = n - 1 To 0 Step -1
    i = InStrRev(f, "\", IIf(j = n - 1, -1, i - 1))
    If i > 0 Then x = i Else Exit For
Next j

TrncFilePath = Mid$(f, x + 1)
End Function

Function RemExtFromPath(f As String) As String
Dim x As Long
x = InStrRev(f, ".")
If x > 0 Then
    RemExtFromPath = Mid$(f, 1, x - 1)
Else
    RemExtFromPath = f
End If
End Function

Function FileNameFromPath(f As String) As String
FileNameFromPath = Mid$(f, InStrRev(f, "\") + 1)
End Function

Function FileFolderFromPath(f As String) As String
FileFolderFromPath = Mid$(f, 1, InStrRev(f, "\"))
End Function

Public Function FileExt(f As String) As String
FileExt = Mid(f, InStrRev(f, ".") + 1)
End Function

Function IsCompiled() As Boolean
Dim b As Boolean
b = False
Debug.Assert IsCompiled_check(b)
IsCompiled = Not b
End Function

Private Function IsCompiled_check(b As Boolean) As Boolean
b = True
IsCompiled_check = True
End Function

Attribute VB_Name = "modFileSys"

Option Explicit
Option Compare Binary
Option Base 0

Public fsoMain As New FileSystemObject

'directory listing stuff=======================================================
'Public Sub debug_testListDirs()
'Dim cFlds As Collection, i As Long
'
'Set cFlds = BuildDirCollection("D:\data")
'
'Debug.Print cFlds.count & " items."
'For i = 1 To cFlds.count
'    Debug.Print cFlds(i)
'Next i
'End Sub

Public Function AddFilesToCollection(sFolder As String, cFiles As Collection) As Collection
On Error GoTo ErrHnd

Dim fsFolder As Folder, fsFile As File, e As String

e = "1"
Set fsFolder = fsoMain.GetFolder(sFolder)

e = "2"
For Each fsFile In fsFolder.Files
    cFiles.Add fsFile.Path
Next

Exit Function
ErrHnd:
Main_Err "error in sub 'AddFilesToCollection';" & e & "."
End Function

Public Sub BuildDirCollection(f As String, cRet As Collection, _
    Optional txtLog As TextBox)

On Error GoTo ErrHnd

Dim fsTopDir As Folder, e As String

e = "1"
If fsoMain.DriveExists(f) Then
    Set fsTopDir = fsoMain.GetDrive(f).RootFolder
ElseIf fsoMain.FolderExists(f) Then
    Set fsTopDir = fsoMain.GetFolder(f)
Else
    MsgBox "error: '" & f & "' does not exist."
    Exit Sub
End If

e = "2"
RecurseSubFolders fsTopDir, cRet, txtLog

e = "3"

Exit Sub
ErrHnd:
If txtLog Is Empty Then
    Main_Err "error in sub 'BuildDirCollection';" & e & "."
Else
    AddToLog txtLog, "error in sub 'BuildDirCollection';" & e & "." & vbNewLine & _
        vbTab & "internal description: " & e & vbNewLine & _
        vbTab & "system description: " & Err.Number & "; " & Err.Description & "." & vbNewLine
End If
End Sub

Private Function RecurseSubFolders(fsFld As Folder, cFlds As Collection, _
    Optional txtLog As TextBox)

On Error GoTo ErrHnd

Dim fsSubFld As Folder, e As String

e = "1"
e = "1-" & fsFld.Path & "-" & fsFld.SubFolders.Count
cFlds.Add fsFld.Path

e = "2"
If Not IsEmpty(fsFld.SubFolders) Then
    e = "3"
    For Each fsSubFld In fsFld.SubFolders
        e = "4"
        RecurseSubFolders fsSubFld, cFlds, txtLog
    Next
End If

e = "5"

Exit Function
ErrHnd:
If txtLog Is Empty Then
    Main_Err "error in sub 'RecurseSubFolders';" & e & "."
Else
    AddToLog txtLog, "error in sub 'RecurseSubFolders';" & e & "." & vbNewLine & _
        vbTab & "internal description: " & e & vbNewLine & _
        vbTab & "system description: " & Err.Number & "; " & Err.Description & "." & vbNewLine
End If
End Function

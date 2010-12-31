Attribute VB_Name = "modConsols"
Option Explicit

Public Const y_ChangeAnimePosEvery As Long = 300 'seconds

Public Const y_MaxConsoles As Long = 10
Public Const y_NextDelBase As Long = 1
Public Const y_NextDelVar As Long = 3
Public Const y_CloseDelBase As Long = 1
Public Const y_CloseDelVar As Long = 5
Public Const y_SnipsMinN As Long = 2
Public Const y_SnipsVarN As Long = 4

Public Const y_CodeLineInt As Long = 20 '50 'ms
Public Const y_CodeChrInt As Long = 30 '50 'ms

Public y_ConsoleCnt As Long
Public y_NextConIn As Long
Public y_NextConI As Long

Public y_AllCodeSnips() As String

Public y_arrCons(0 To y_MaxConsoles - 1) As frmConsole

Public Sub y_Init(Frm As Form)
Dim i As Long

If Not Frm Is Nothing Then
    Frm.SetStat "loading snipets..."
End If

y_ConsoleCnt = 0
y_NextConI = 0
y_NextConIn = 0
y_LoadCodeSnips Frm

For i = 0 To UBound(y_arrCons)
    Set y_arrCons(i) = New frmConsole
Next i
End Sub

Private Sub y_LoadCodeSnips(Optional Frm As Form)
Dim cCodeFiles As New Collection, i As Long, arrF() As String, _
    tsFile As TextStream

'build list of files
AddFilesToCollection App.Path & "\" & dir_PseudoCode, cCodeFiles
ReDim arrF(0)
For i = 1 To cCodeFiles.Count
    If FileExt(cCodeFiles(i)) = "txt" Then
        If i - 1 > UBound(arrF) Then ReDim Preserve arrF(0 To UBound(arrF) + 1)
        arrF(UBound(arrF)) = cCodeFiles(i)
    End If
Next i

'load these files
ReDim y_AllCodeSnips(0)
For i = 0 To UBound(arrF)
    If Not Frm Is Nothing Then
        Frm.SetStat "loading snipet " & i & " of " & UBound(arrF) & "..."
    End If
    
    Set tsFile = fsoMain.OpenTextFile(arrF(i))
    If i > UBound(y_AllCodeSnips) Then ReDim Preserve y_AllCodeSnips(0 To UBound(y_AllCodeSnips) + 1)
    y_AllCodeSnips(UBound(y_AllCodeSnips)) = tsFile.ReadAll
    tsFile.Close
    Set tsFile = Nothing
Next i
End Sub

Public Sub y_CheckToSpawnNew()
If y_ConsoleCnt >= y_MaxConsoles Then Exit Sub

If y_NextConIn > y_NextConI Then
    y_NextConI = y_NextConI + 1
    Exit Sub
End If

y_SpawnNew

y_NextConI = 0
Randomize Timer
y_NextConIn = (Rnd() * y_NextDelVar) + y_NextDelBase
End Sub

Private Function y_FindUnusedFrm() As frmConsole
Dim i As Long
Set y_FindUnusedFrm = Nothing

If y_ConsoleCnt >= y_MaxConsoles Then Exit Function

For i = 0 To UBound(y_arrCons)
    If Not y_arrCons(i).m_Active Then
        Set y_FindUnusedFrm = y_arrCons(i)
        Exit For
    End If
Next i
End Function

Public Sub y_SpawnNew()
Dim Frm As frmConsole, lL As Long, lT As Long, lMon As Long, _
    lScrLeft As Long, lScrTop As Long, lScrWidth As Long, lScrHeight As Long

'create
Set Frm = y_FindUnusedFrm
If Frm Is Nothing Then Exit Sub

'set up die timer (starts once code has finished)
Randomize Timer
Frm.Init Int(Rnd() * y_CloseDelVar) + y_CloseDelBase

'choose a monitor
If m_cM.MonitorCount > 1 Then
    lMon = Int(Rnd() * m_cM.MonitorCount) + 1
Else
    lMon = 1
End If

'get screen size
lScrLeft = m_cM.Monitor(lMon).WorkLeft * Screen.TwipsPerPixelX
lScrTop = m_cM.Monitor(lMon).WorkTop * Screen.TwipsPerPixelY
lScrWidth = m_cM.Monitor(lMon).WorkWidth * Screen.TwipsPerPixelX
lScrHeight = m_cM.Monitor(lMon).WorkHeight * Screen.TwipsPerPixelY

'set form position, not overlapping frmMain
lL = lScrLeft + Int(Rnd() * (lScrWidth - Frm.Width))
If prf_ShowAnim And lMon = 1 Then
    lT = -1
    Select Case frmMain.m_Corner
        Case 0 'top left
            If lL < frmMain.Left + frmMain.Width Then
                lT = frmMain.Top + frmMain.Height + Int(Rnd() * (lScrHeight - frmMain.Top - frmMain.Height - Frm.Height))
            End If
        
        Case 1 'top right
            If lL + Frm.Width > frmMain.Left Then
                lT = frmMain.Top + frmMain.Height + Int(Rnd() * (lScrHeight - frmMain.Top - frmMain.Height - Frm.Height))
            End If
        
        Case 2 'btm left
            If lL < frmMain.Left + frmMain.Width Then
                lT = lScrTop + Int(Rnd() * (frmMain.Top - Frm.Height))
            End If
        
        Case 3 'btm right
            If lL + Frm.Width > frmMain.Left Then
                lT = lScrTop + Int(Rnd() * (frmMain.Top - Frm.Height))
            End If
        
    End Select
    
    If lT < 0 Then lT = Int(Rnd() * (lScrHeight - Frm.Height))
Else
    lT = lScrTop + Int(Rnd() * (lScrHeight - Frm.Height))
End If

Frm.Left = lL
Frm.Top = lT

Frm.Start
y_ConsoleCnt = y_ConsoleCnt + 1
End Sub

Public Sub y_ConsoleAboutToDie(Frm As frmConsole)
If y_ConsoleCnt = 1 Then
    y_NextConI = 0
    y_NextConIn = 0
End If
End Sub

Public Sub y_ConsoleDie(Frm As frmConsole)
'Unload Frm
'Set Frm = Nothing
Frm.Unstart
y_ConsoleCnt = y_ConsoleCnt - 1
End Sub

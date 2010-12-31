Attribute VB_Name = "modMain"
Option Explicit

Public Const dir_PseudoCode As String = "pseudo-code"
Public Const file_YukiAnimGif As String = "yuki420.gif"
Public Const file_Ini As String = "yuki.ini"

Public prf_ProcPri As Boolean
Public prf_ShowAnim As Boolean
Public prf_ConsoleText As Long
Public prf_ConsoleBack As Long

Public m_cM As New cMonitors

Private Const APP_NAME = "yuki.scr"

Sub Main()
CheckPrevInstance
LoadSettings

Select Case Mid(UCase$(Trim$(Command$)), 1, 2)
    Case "/C" 'Configurations mode called
        frmPref.Show
    Case "/A" 'Password protect dialog
        MsgBox "Password Protection not available with this screen saver."
    Case "/P" 'Preview mode
        End
    Case Else
        main_Start
End Select
End Sub

Private Sub CheckPrevInstance()
If App.PrevInstance Then main_End
If FindWindow(vbNullString, APP_NAME) Then main_End
End Sub

Private Sub main_Start()
If prf_ProcPri Then SetMyPri ABOVE_NORMAL_PRIORITY_CLASS

If prf_ShowAnim Then
    Load frmMain
    frmMain.Caption = APP_NAME
    frmMain.SetFrmPos
    frmMain.Show
    frmMain.tmrShow.Enabled = True
End If
y_Init frmMain
If prf_ShowAnim Then frmMain.Anim_Act
frmMain.Anim_Start
End Sub

Public Sub LoadSettings()
Dim f As String, a
f = App.Path & "\" & file_Ini

a = GetFromIniEx("gen", "procpri", "0", f)  'default to no
prf_ProcPri = IIf(a = "1", True, False)

a = GetFromIniEx("gen", "showanim", "1", f) 'default to yes
prf_ShowAnim = IIf(a = "1", True, False)

a = GetFromIniEx("console", "textcolour", "16777215", f)
prf_ConsoleText = Val(a)

a = GetFromIniEx("console", "backcolour", "0", f)
prf_ConsoleBack = Val(a)
End Sub

Public Sub SaveSettings()
Dim f As String
f = App.Path & "\" & file_Ini

WriteToIni "gen", "procpri", IIf(prf_ProcPri, "1", "0"), f
WriteToIni "gen", "showanim", IIf(prf_ShowAnim, "1", "0"), f

WriteToIni "console", "textcolour", Trim$(Str$(prf_ConsoleText)), f
WriteToIni "console", "backcolour", Trim$(Str$(prf_ConsoleBack)), f
End Sub

Public Sub main_End()
Dim f As Form
For Each f In Forms
    Unload f
    Set f = Nothing
Next f
End Sub

'types:
'0-critical (an error that is my fault.)
'1-not critical (an error that is not my fault.  e.g. file not found.)
Sub Main_Err(e As String, Optional iType As Integer = 0)
Dim a As String, sSysErr As String
sSysErr = Err.Number & "; " & Err.Description & "."

On Error GoTo Main_Err_err

Select Case iType
    Case 0
        a = "//yuki.scr has encountered a critical error." & vbNewLine & vbNewLine & _
            "internal description: " & vbNewLine & _
            e & vbNewLine & vbNewLine & _
            "system description: " & vbNewLine & _
            sSysErr & vbNewLine & vbNewLine & _
            "version: " & App.Major & "." & App.Minor & "(" & App.Revision & ")" & vbNewLine & vbNewLine & _
            "for supprt please visit aefaradien.net/yuki."
    
    Case 1
        a = "//yuki.scr has encountered a non-critical error." & vbNewLine & vbNewLine & _
            "internal description: " & vbNewLine & _
            e & vbNewLine & vbNewLine & _
            "if you need help solving this error, please visit aefaradien.net/yuki."
        
    
End Select

MsgBox a

Exit Sub
Main_Err_err:
MsgBox "unable to show error window.  this is not a good sign." & vbNewLine & a
End Sub

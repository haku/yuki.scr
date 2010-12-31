VERSION 5.00
Begin VB.Form frmConsole 
   BackColor       =   &H00000000&
   Caption         =   "cmd.exe"
   ClientHeight    =   3680
   ClientLeft      =   40
   ClientTop       =   280
   ClientWidth     =   6920
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmConsole.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3680
   ScaleWidth      =   6920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTyping 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer tmrCode 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtCmd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2650
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   6850
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   0
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_Active As Boolean

Private m_lLife As Long 'life in seconds
Private m_lAge As Long 'age in seconds
Private m_CodeLines() As String
Private m_CodeLinesI As Long
Private m_CodeTypingI As Long

Public Sub Init(lLife As Long)
m_lLife = lLife
tmrHide_Timer
LoadCode
End Sub

Public Sub Start()
Show
tmrCode.Enabled = True
m_Active = True
End Sub

Public Sub Unstart()
Hide

On Error Resume Next
Dim t As Timer
For Each t In Me.Controls
    t.Enabled = False
Next t
On Error GoTo 0

m_lLife = 0
m_lAge = 0
ReDim m_CodeLines(0)
m_CodeLinesI = 0
m_CodeTypingI = 0
txtCmd.Text = ""
m_Active = False
End Sub

Private Sub LoadCode()
Dim i As Long, x As Long, arr() As String

ReDim m_CodeLines(0 To 0)
m_CodeLines(0) = ""
m_CodeLinesI = -1

'pick some snips at random
Randomize Timer
x = Int(Rnd() * y_SnipsVarN) + y_SnipsMinN
ReDim arrF(0 To x)
For i = 0 To UBound(arrF)
    x = Int(Rnd() * (UBound(y_AllCodeSnips) + 1))
    
    arr = Split(y_AllCodeSnips(x), vbNewLine)
    For x = 0 To UBound(arr)
        If Len(m_CodeLines(UBound(m_CodeLines))) > 0 Then ReDim Preserve m_CodeLines(0 To UBound(m_CodeLines) + 1)
        m_CodeLines(UBound(m_CodeLines)) = arr(x)
    Next x
Next i
End Sub

Private Sub ConsoleAdd(s As String)
If Len(txtCmd.Text) + Len(s) >= 65535 Then
    txtCmd.Text = ""
End If

txtCmd.SelStart = Len(txtCmd.Text)
txtCmd.SelText = s
txtCmd.SelStart = Len(txtCmd.Text)
End Sub

Private Sub Form_Activate()
KeepOnTop Me, True
End Sub

Private Sub Form_Load()
m_Active = False
m_lAge = 0

txtCmd.ForeColor = prf_ConsoleText
txtCmd.BackColor = prf_ConsoleBack

tmrCode.Interval = y_CodeLineInt
tmrTyping.Interval = y_CodeChrInt
End Sub

Private Sub Form_Resize()
txtCmd.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unstart
End Sub

Private Sub tmrCode_Timer()
m_CodeLinesI = m_CodeLinesI + 1

If m_CodeLinesI > UBound(m_CodeLines) Then
    tmrCode.Enabled = False
    tmrHide.Enabled = True
ElseIf Mid(m_CodeLines(m_CodeLinesI), 1, 1) = "." Or _
    Mid(m_CodeLines(m_CodeLinesI), 1, 1) = "," Then
    tmrCode.Enabled = False
    m_CodeLines(m_CodeLinesI) = Mid(m_CodeLines(m_CodeLinesI), 2)
    m_CodeTypingI = 0
    If Len(txtCmd.Text) > 0 Then
        ConsoleAdd vbNewLine
    End If
    tmrTyping.Enabled = True
    If Mid(m_CodeLines(m_CodeLinesI), 1, 1) = "." Then Show
Else
    If Len(txtCmd.Text) > 0 Then
        ConsoleAdd vbNewLine & m_CodeLines(m_CodeLinesI)
    Else
        ConsoleAdd m_CodeLines(m_CodeLinesI)
    End If
End If
End Sub

Private Sub tmrTyping_Timer()
ConsoleAdd Mid(m_CodeLines(m_CodeLinesI), m_CodeTypingI + 1, 1)

m_CodeTypingI = m_CodeTypingI + 1
If m_CodeTypingI > Len(m_CodeLines(m_CodeLinesI)) Then
    tmrTyping.Enabled = False
    tmrCode.Enabled = True
End If
End Sub

Private Sub tmrHide_Timer()
If Not tmrHide.Enabled Then Exit Sub

m_lAge = m_lAge + 1
If m_lAge > m_lLife Then
    y_ConsoleDie Me
    tmrHide.Enabled = False
ElseIf m_lLife - m_lAge < 2 Then '2 second warning so another can be spawned just before this one closes
    y_ConsoleAboutToDie Me
End If
End Sub

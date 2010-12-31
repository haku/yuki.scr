VERSION 5.00
Begin VB.Form frmPref 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "yuki.scr preferences"
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4940
   Icon            =   "frmPref.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   250
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   490
   End
   Begin VB.PictureBox picColour 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   250
      Index           =   1
      Left            =   1800
      ScaleHeight     =   210
      ScaleWidth      =   450
      TabIndex        =   9
      Top             =   1560
      Width           =   490
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   250
      Index           =   0
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   490
   End
   Begin VB.PictureBox picColour 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   250
      Index           =   0
      Left            =   1800
      ScaleHeight     =   210
      ScaleWidth      =   450
      TabIndex        =   7
      Top             =   1200
      Width           =   490
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   300
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Width           =   970
   End
   Begin VB.CheckBox chkAnimWin 
      Caption         =   "show animation window."
      Height          =   250
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4690
   End
   Begin VB.CheckBox chkChangePri 
      Caption         =   "change process priority to ""above normal""."
      Height          =   250
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4690
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "cancel"
      Height          =   300
      Left            =   3840
      TabIndex        =   0
      Top             =   2280
      Width           =   970
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "console back colour:"
      Height          =   200
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1520
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "console text colour:"
      Height          =   200
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      Caption         =   "version="
      Height          =   200
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   640
   End
End
Attribute VB_Name = "frmPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
main_End
End Sub

Private Sub cmdColour_Click(Index As Integer)
Dim c As Long
c = ShowColorDlg(hwnd, picColour(Index).BackColor)
If c > -1 Then picColour(Index).BackColor = c
End Sub

Private Sub cmdOk_Click()
prf_ProcPri = IIf(chkChangePri.Value = 1, True, False)
prf_ShowAnim = IIf(chkAnimWin.Value = 1, True, False)

prf_ConsoleText = picColour(0).BackColor
prf_ConsoleBack = picColour(1).BackColor

SaveSettings
cmdCancel_Click
End Sub

Private Sub Form_Load()
lblVer.Caption = lblVer.Caption & App.Major & "." & App.Minor & " (build " & App.Revision & ")."

chkChangePri.Value = IIf(prf_ProcPri, 1, 0)
chkAnimWin.Value = IIf(prf_ShowAnim, 1, 0)

picColour(0).BackColor = prf_ConsoleText
picColour(1).BackColor = prf_ConsoleBack
End Sub

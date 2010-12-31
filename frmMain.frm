VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "yuki.scr"
   ClientHeight    =   2080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2080
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrScr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   1560
   End
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Left            =   840
      Top             =   1560
   End
   Begin VB.PictureBox picAnim 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   490
      Left            =   0
      ScaleHeight     =   490
      ScaleWidth      =   970
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   970
   End
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   1560
   End
   Begin VB.Timer tmrEnd 
      Interval        =   500
      Left            =   120
      Top             =   1560
   End
   Begin VB.Label lblLoading 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "loading..."
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
      Height          =   160
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type UcsFrameInfo
    dibFrame As cDIBSection
    nDelay As Long
End Type

Private m_GifRead As clsGifReader
Private m_GifRender As clsBmpRenderer
Private m_GifFrames() As UcsFrameInfo
Private m_GifFramecnt As Long
Private m_GifCurrentframe As Long

Private m_Mouse As POINTAPI

Public m_Corner As Long, m_ChangeCornerIn As Long

Private Sub Anim_Init()
Set m_GifRead = New clsGifReader
Set m_GifRender = New clsBmpRenderer
Set picAnim.Picture = Nothing

m_GifRead.Init App.Path & "\" & file_YukiAnimGif
m_GifRender.Init m_GifRead

If m_GifRead.MoveLast() Then
    m_GifFramecnt = m_GifRead.FrameIndex + 1
    picAnim.Move 0, 0, _
        m_GifRead.ImageWidth * Screen.TwipsPerPixelX, _
        m_GifRead.ImageHeight * Screen.TwipsPerPixelY
End If
End Sub

Public Sub Anim_Act()
Dim i As Long
ReDim m_GifFrames(-1 To -1)

DoEvents

If m_GifFramecnt > 0 Then
    ReDim m_GifFrames(1 To m_GifFramecnt)
    
    If m_GifRender.MoveFirst() Then
        i = 0
        Do While True
            If Not m_GifRender.MoveNext Then Exit Do
            i = i + 1
            
            SetStat "loading frame " & i & " of " & m_GifFramecnt & "..."
            
            'Set m_GifFrames(i).oPic = m_GifRender.Image
            Set m_GifFrames(i).dibFrame = New cDIBSection
            m_GifFrames(i).dibFrame.CreateFromPicture m_GifRender.Image
            m_GifFrames(i).nDelay = m_GifRender.Reader.DelayTime
            'DoEvents
            
            If m_GifRender.EOF Then Exit Do
        Loop
        
        picAnim.Visible = True
        m_GifCurrentframe = 1
        tmrAnim_Timer
        tmrAnim.Enabled = True
    End If
End If
End Sub

Public Sub Anim_Start()
tmrScr.Enabled = True
End Sub

Public Sub SetFrmPos(Optional lCorner As Long = 3)
Dim lL As Long, lT As Long, lW As Long, lH As Long, lMargin As Long, _
    lScrLeft As Long, lScrTop As Long, lScrWidth As Long, lScrHeight As Long

lW = picAnim.Width
lH = picAnim.Height
lMargin = lH * 0.2

'place on first monitor
lScrLeft = m_cM.Monitor(1).WorkLeft * Screen.TwipsPerPixelX
lScrTop = m_cM.Monitor(1).WorkTop * Screen.TwipsPerPixelY
lScrWidth = m_cM.Monitor(1).WorkWidth * Screen.TwipsPerPixelX
lScrHeight = m_cM.Monitor(1).WorkHeight * Screen.TwipsPerPixelY

Select Case lCorner
    Case 0 'top left
        lL = lScrLeft + lMargin
        lT = lScrTop + lMargin
    
    Case 1 'top right
        lL = lScrLeft + lScrWidth - lW - lMargin
        lT = lScrTop + lMargin
    
    Case 2 'btm left
        lL = lScrLeft + lMargin
        lT = lScrTop + lScrHeight - lH - lMargin
    
    Case 3 'btm right
        lL = lScrLeft + lScrWidth - lW - lMargin
        lT = lScrTop + lScrHeight - lH - lMargin
    
End Select

frmMain.Move lL, lT, lW, lH

m_Corner = lCorner
m_ChangeCornerIn = y_ChangeAnimePosEvery
End Sub

Public Sub SetStat(s As String)
lblLoading.Caption = "YUKI.N>" & s & " _"
lblLoading.Refresh
End Sub

Private Sub Form_Activate()
KeepOnTop Me, True
End Sub

Private Sub Form_Load()
lblLoading.Move 200, 200
Anim_Init
GetCursorPos m_Mouse
End Sub

Private Sub picAnim_DblClick()
main_End
End Sub

Private Sub picAnim_Paint()
If m_GifCurrentframe < 1 Then Exit Sub
With m_GifFrames(m_GifCurrentframe).dibFrame
    .PaintPicture picAnim.hDC, 0, 0, .Width, .Height, 0, 0, vbSrcCopy
End With
End Sub

Private Sub tmrAnim_Timer()
Dim lDelay As Long

If tmrAnim.Enabled Then m_GifCurrentframe = (m_GifCurrentframe Mod m_GifFramecnt) + 1

lDelay = IIf(m_GifFrames(m_GifCurrentframe).nDelay < 8, 80, m_GifFrames(m_GifCurrentframe).nDelay * 10)
picAnim_Paint
tmrAnim.Interval = lDelay
If tmrAnim.Enabled Then
    tmrAnim.Enabled = False
    tmrAnim.Enabled = True
End If
End Sub

Private Sub tmrEnd_Timer()
'checks for mouse movement
Dim pt As POINTAPI, lDist As Long
GetCursorPos pt
lDist = Sqr(Abs(m_Mouse.x - pt.x) + Abs(m_Mouse.y - pt.y))
If lDist >= 5 Then
    main_End
End If
End Sub

Private Sub tmrScr_Timer()
y_CheckToSpawnNew

m_ChangeCornerIn = m_ChangeCornerIn - 1
If m_ChangeCornerIn <= 0 Then
    Randomize Timer
    SetFrmPos Int(Rnd() * 4)
End If
End Sub

Private Sub tmrShow_Timer()
'keeps this window visible after launch
WindowState = vbNormal
Show
tmrShow.Enabled = False
End Sub

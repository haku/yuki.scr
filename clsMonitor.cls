VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Private Type RECT
   Left As Long
   Top As Long
   right As Long
   bottom As Long
End Type

Private Const CCHDEVICENAME = 32

Private Type MONITORINFOEXA
   cbSize As Long
   rcMonitor As RECT
   rcWork As RECT
   dwFlags As Long
   b(0 To CCHDEVICENAME - 1) As Byte
End Type

Private Type MONITORINFOEXW
   cbSize As Long
   rcMonitor As RECT
   rcWork As RECT
   dwFlags As Long
   b(0 To CCHDEVICENAME * 2 - 1) As Byte
End Type

Private Declare Function GetMonitorInfoA Lib "User32" ( _
      ByVal hMonitor As Long, _
      lpmi As MONITORINFOEXA _
   ) As Long
Private Declare Function GetMonitorInfoW Lib "User32" ( _
      ByVal hMonitor As Long, _
      lpmi As MONITORINFOEXW _
   ) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const MONITORINFOF_PRIMARY = &H1

Private m_hMonitor As Long
Private m_sName As String
Private m_rcMonitor As RECT
Private m_rcWork As RECT
Private m_bIsPrimary As Boolean

Public Property Get IsPrimary() As Boolean
   IsPrimary = m_bIsPrimary
End Property

Public Property Get hMonitor() As Long
   hMonitor = m_hMonitor
End Property
Public Property Get Name() As String
   Name = m_sName
End Property
Public Property Get Left() As Long
   Left = m_rcMonitor.Left
End Property
Public Property Get Top() As Long
   Top = m_rcMonitor.Top
End Property
Public Property Get Width() As Long
   Width = m_rcMonitor.right - m_rcMonitor.Left
End Property
Public Property Get Height() As Long
   Height = m_rcMonitor.bottom - m_rcMonitor.Top
End Property
Public Property Get WorkLeft() As Long
   WorkLeft = m_rcWork.Left
End Property
Public Property Get WorkTop() As Long
   WorkTop = m_rcWork.Top
End Property
Public Property Get WorkWidth() As Long
   WorkWidth = m_rcWork.right - m_rcMonitor.Left
End Property
Public Property Get WorkHeight() As Long
   WorkHeight = m_rcWork.bottom - m_rcMonitor.Top
End Property

Friend Sub fInit( _
      ByVal hMonitor As Long _
   )
Dim iPos As Long
Dim sName As String
   m_hMonitor = hMonitor
   If (IsNt) Then
      Dim tMIW As MONITORINFOEXW
      tMIW.cbSize = Len(tMIW)
      GetMonitorInfoW hMonitor, tMIW
      With tMIW
         LSet m_rcMonitor = .rcMonitor
         LSet m_rcWork = .rcWork
         m_bIsPrimary = ((.dwFlags And MONITORINFOF_PRIMARY) = MONITORINFOF_PRIMARY)
         sName = .b
         iPos = InStr(sName, vbNullChar)
      End With
   Else
      Dim tMIA As MONITORINFOEXA
      tMIA.cbSize = Len(tMIA)
      GetMonitorInfoA hMonitor, tMIA
      With tMIA
         LSet m_rcMonitor = .rcMonitor
         LSet m_rcWork = .rcWork
         m_bIsPrimary = ((.dwFlags And MONITORINFOF_PRIMARY) = MONITORINFOF_PRIMARY)
         sName = StrConv(.b, vbUnicode)
      End With
   End If
   iPos = InStr(sName, vbNullChar)
   If (iPos > 0) Then
      m_sName = Left(sName, iPos - 1)
   Else
      m_sName = sName
   End If
   
End Sub

'the more you ignore me/
'the closer i get/
'you're wasting your time...
'beware/
'I bear more grudges/
'than lonely high court judges/
'when you sleep/
'I will creep/
'into your thoughts/
'like a bad debt/
'that you can't pay/
'take the easy way/
'and give in/
'and let me in...


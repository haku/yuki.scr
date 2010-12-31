VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMonitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79
Private Const SM_CMONITORS = 80
Private Const SM_SAMEDISPLAYFORMAT = 81

Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long

Private Declare Function MonitorFromWindow Lib "User32" (ByVal hWND As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromPoint Lib "User32" (pt As POINTAPI, ByVal dwFlags As Long) As Long
Private Const MONITOR_DEFAULTTONEAREST = 0

Private m_iCount As Long
Private m_cM() As cMonitor

Public Property Get AllMonitorsSame() As Long
   AllMonitorsSame = GetSystemMetrics(SM_SAMEDISPLAYFORMAT)
End Property

Public Property Get MonitorForPoint(ByVal x As Long, ByVal y As Long) As cMonitor
Dim hMon As Long
Dim tP As POINTAPI
   tP.x = x
   tP.y = y
   hMon = MonitorFromPoint(tP, MONITOR_DEFAULTTONEAREST)
   If Not (hMon = 0) Then
      Dim cM As cMonitor
      Set cM = New cMonitor
      cM.fInit hMon
      Set MonitorForPoint = cM
   End If
End Property

Public Property Get MonitorForWindow(ByVal hWND As Long) As cMonitor
Dim hMon As Long
   hMon = MonitorFromWindow(hWND, MONITOR_DEFAULTTONEAREST)
   If Not (hMon = 0) Then
      Dim cM As cMonitor
      Set cM = New cMonitor
      cM.fInit hMon
      Set MonitorForWindow = cM
   End If
End Property

Public Property Get VirtualScreenLeft() As Long
Dim lRet As Long
Dim i As Long
   lRet = m_cM(1).Left
   For i = 2 To m_iCount
      If (m_cM(i).Left < lRet) Then
         lRet = m_cM(i).Left
      End If
   Next i
   VirtualScreenLeft = lRet
End Property
Public Property Get VirtualScreenTop() As Long
Dim lRet As Long
Dim i As Long
   lRet = m_cM(1).Top
   For i = 2 To m_iCount
      If (m_cM(i).Top < lRet) Then
         lRet = m_cM(i).Top
      End If
   Next i
   VirtualScreenTop = lRet
End Property
Public Property Get VirtualScreenWidth() As Long
   VirtualScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
End Property
Public Property Get VirtualScreenHeight() As Long
   VirtualScreenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
End Property
Public Property Get DisplayMonitorCount() As Long
   DisplayMonitorCount = GetSystemMetrics(SM_CMONITORS)
End Property
Public Property Get MonitorCount() As Long
   MonitorCount = m_iCount
End Property
Public Property Get Monitor(ByVal index As Long) As cMonitor
   Set Monitor = m_cM(index)
End Property

Friend Sub fAddMonitor( _
      ByVal hMonitor As Long _
   )
   
   ' this sub will not actually provide an
   ' extra monitor to your system if it isn't
   ' already there.  yet.
   
   m_iCount = m_iCount + 1
   ReDim Preserve m_cM(1 To m_iCount) As cMonitor
   Set m_cM(m_iCount) = New cMonitor
   m_cM(m_iCount).fInit hMonitor
     
End Sub

Public Sub Refresh()
   m_iCount = 0
   Erase m_cM
   EnumMonitors Me
End Sub

Private Sub Class_Initialize()
   Refresh
End Sub

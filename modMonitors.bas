Attribute VB_Name = "mMonitors"

Option Explicit
Option Compare Binary
Option Base 0

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function EnumDisplayMonitors Lib "User32" ( _
      ByVal hDC As Long, _
      lprcClip As Any, _
      ByVal lpfnEnum As Long, _
      ByVal dwData As Long _
   ) As Long
Private m_cM As cMonitors

Private Function MonitorEnumProc( _
      ByVal hMonitor As Long, _
      ByVal hDCMonitor As Long, _
      ByVal lprcMonitor As Long, _
      ByVal dwData As Long _
   ) As Long
   m_cM.fAddMonitor hMonitor
   MonitorEnumProc = 1
End Function

Public Sub EnumMonitors(cM As cMonitors)
   Set m_cM = cM
   EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, 0
End Sub

Public Function IsNt() As Boolean
Dim lVer As Long
   lVer = GetVersion()
   IsNt = ((lVer And &H80000000) = 0)
End Function


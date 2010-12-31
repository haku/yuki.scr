Attribute VB_Name = "modPri"
Option Explicit

Public Const THREAD_BASE_PRIORITY_IDLE = -15
Public Const THREAD_BASE_PRIORITY_LOWRT = 15
Public Const THREAD_BASE_PRIORITY_MIN = -2
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
Public Const THREAD_PRIORITY_NORMAL = 0
Public Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT

Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100

Public Const BELOW_NORMAL_PRIORITY_CLASS = &H4000
Public Const ABOVE_NORMAL_PRIORITY_CLASS = 32768

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Function SetMyPri(lPri As Long, Optional lThr As Long = THREAD_PRIORITY_NORMAL)
Dim hThread As Long, hProcess As Long

hThread = GetCurrentThread
hProcess = GetCurrentProcess

SetThreadPriority hThread, lThr
SetPriorityClass hProcess, lPri

'print some results
'Debug.Print "Current Thread Priority:" + Str$(GetThreadPriority(hThread))
'Debug.Print "Current Priority Class:" + Str$(GetPriorityClass(hProcess))
End Function

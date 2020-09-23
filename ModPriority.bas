Attribute VB_Name = "ModPriority"
'******************************************************************************************************
'* WARNING It Is Highly Recommended Not Too Alter NextPad's Default Priority Level....               *
'* It Can Cause Unpredictable Results On Windows 2000 , ME , NT , 95 ,                               *
'* It hasnt been Tested On these Os's Yet so i Suggest Leaving This Settings Alone                   *
'*****************************************************************************************************

Public Const Priority_RealTime = "realtime"
Public Const Priority_Idle = "idle"
Public Const Priority_Normal = "normal"
Public Const Priority_Highest = "highest"

Const THREAD_BASE_PRIORITY_IDLE = -15
Const THREAD_BASE_PRIORITY_LOWRT = 15
Const THREAD_BASE_PRIORITY_MIN = -2
Const THREAD_BASE_PRIORITY_MAX = 2
Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
Const THREAD_PRIORITY_NORMAL = 0
Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Const HIGH_PRIORITY_CLASS = &H80
Const IDLE_PRIORITY_CLASS = &H40
Const NORMAL_PRIORITY_CLASS = &H20
Const REALTIME_PRIORITY_CLASS = &H100
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Function SetPriority(nPriority As Integer)
 
 Dim lThread As Long
  lThread = GetCurrentThread
   Select Case nPriority
     Case 0
       SetThreadPriority lThread, THREAD_PRIORITY_IDLE
        SetThreadPriority lThread, IDLE_PRIORITY_CLASS
     Case 1
       SetThreadPriority lThread, THREAD_PRIORITY_NORMAL
        SetPriorityClass lThread, NORMAL_PRIORITY_CLASS
     Case 2
       SetThreadPriority lThread, THREAD_PRIORITY_HIGHEST
         SetPriorityClass lThread, HIGH_PRIORITY_CLASS
     Case 3
       SetThreadPriority lThread, REALTIME_PRIORITY_CLASS
        SetPriorityClass lThread, REALTIME_PRIORITY_CLASS
     Case Else
        Exit Function
   End Select
End Function

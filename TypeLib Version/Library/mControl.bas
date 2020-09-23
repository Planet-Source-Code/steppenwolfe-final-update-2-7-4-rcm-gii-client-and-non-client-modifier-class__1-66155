Attribute VB_Name = "mControl"
Option Explicit

Private Const THREAD_SET_INFORMATION     As Long = &H20

Public Enum ePriority
    Thread_Idle = -15
    Thread_LowRT = 15
    Thread_Minimum = -2
    Thread_Normal = 0
    Thread_Maximum = 2
End Enum

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, _
                                                           ByVal nPriority As Long) As Long


Private Declare Function OpenThread Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, _
                                                        ByVal bInheritHandle As Boolean, _
                                                        ByVal dwThreadId As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Function ThreadAccelerate(ByVal lThread As Long, _
                                  eLevel As ePriority) As Boolean
'/* alter thread priority

Dim lPriority As Long
Dim lReturn   As Long
Dim lHandle   As Long

On Error GoTo Handler

    If lThread <= 0 Then
        GoTo Handler
    End If
    lPriority = eLevel
    lHandle = OpenThread(THREAD_SET_INFORMATION, False, lThread)
    lReturn = SetThreadPriority(lHandle, lPriority)
    '/* success
    If Not lReturn = 0 Then
        ThreadAccelerate = True
    Else
        GoTo Handler
    End If
    
'/* cleanup
CloseHandle lHandle
On Error GoTo 0
Exit Function

Handler:
    If Not lHandle = 0 Then
        CloseHandle lHandle
    End If
    
End Function

Public Sub Error_Data(ByVal sMod As String, _
                      ByVal sProc As String, _
                      ByVal lErrNum As Long, _
                      ByVal sErrDesc As String, _
                      Optional ByVal lLine As Long)

    Open App.Path & "\errlog.log" For Append As #1
    Print #1, "Module: " & sMod & " Proceedure: " & sProc & "Error# " & _
    CStr(lErrNum) & " Desc: " & sErrDesc & " Line: " & CStr(lLine)
    Close #1
    
End Sub

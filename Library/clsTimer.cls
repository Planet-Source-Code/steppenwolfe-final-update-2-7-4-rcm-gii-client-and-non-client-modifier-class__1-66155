VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'VTable:
'  =Name=         =Type=                        =Offset=
'   Item           INVOKE_PROPERTYGET            28
'   Item           INVOKE_PROPERTYPUT            32
'   Item           INVOKE_PROPERTYPUTREF         36
'   Interval       INVOKE_PROPERTYGET            40
'   Interval       INVOKE_PROPERTYPUT            44
'   PulseTimer     INVOKE_FUNC                   48
' User can attach any Variant data they want to the timer
Private m_Item         As Variant

Public Event ThatTime()
' SubTimer is independent of VBCore, so it hard codes error handling
Public Enum EErrorTimer
    eeBaseTimer = 13650     ' CTimer
    eeTooManyTimers         ' No more than 10 timers allowed per class
    eeCantCreateTimer       ' Can't create system timer
End Enum

Private iInterval   As Long
Private id          As Long

Private Sub Class_Terminate()
    Interval = 0
End Sub

Friend Sub ErrRaise(ByVal e As Long)

Dim sText   As String
Dim sSource As String

    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
        Case eeTooManyTimers
            sText = "No more than 10 timers allowed per class"
        Case eeCantCreateTimer
            sText = "Can't create system timer"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else
    ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
    
End Sub

Public Property Get Item() As Variant
    Item = m_Item
End Property

Public Property Let Item(PropVal As Variant)
    m_Item = PropVal
End Property

Public Property Get Interval() As Long
    Interval = iInterval
End Property

' Can't just change interval--you must kill timer and start a new one
Public Property Let Interval(ByVal iIntervalA As Long)

Dim f As Boolean

    If iIntervalA > 0 Then
        ' Don't mess with it if interval is the same
        If iInterval = iIntervalA Then
            Exit Property
        End If
        ' Must destroy any existing timer to change interval
        If iInterval Then
            f = TimerDestroy(Me)
            Debug.Assert f      ' Shouldn't fail
        End If
        ' Create new timer with new interval
        iInterval = iIntervalA
        If Not TimerCreate(Me) Then
            ErrRaise eeCantCreateTimer
        End If
    Else
        If iInterval > 0 Then
            iInterval = 0
            f = TimerDestroy(Me)
            Debug.Assert f      ' Shouldn't fail
        End If
    End If

End Property

' Must be public so that Timer object can't terminate while client's ThatTime
' event is being processed--Friend wouldn't prevent this disaster
Public Sub PulseTimer()
    RaiseEvent ThatTime
End Sub

Friend Property Get TimerID() As Long
    TimerID = id
End Property

Friend Property Let TimerID(ByVal idA As Long)
    id = idA
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Timer control"
Option Explicit

Option Explicit

'Author:    Flavio Morgado
'Book :     Programming Excel with VBA: A Practical Real-World Guide (Chapter 10)
'Edit :     @thaipv #30/09/2019#

Public Event Timer()

#If VBA7 Then
    Private lngTimerID  As LongPtr
#Else
    Private lngTimerID  As Long
#End If

Private lngInterval     As Long
Private isEnabled       As Boolean

'// Some properties
Friend Property Let Interval(ByVal lngValue As Long)
    If lngValue > 0 And lngValue <> lngInterval Then lngInterval = lngValue
End Property
Friend Property Get Interval() As Long
    Interval = lngInterval
End Property
Friend Property Get Enabled() As Boolean
    Enabled = isEnabled
End Property
Friend Property Let Enabled(ByVal isValue As Boolean)
    If isValue Then
        If 0 = lngTimerID And Me.Interval > 0 Then
            lngTimerID = SetTimer(Application.hWnd, ObjPtr(Me), _
                                  Me.Interval, GetPtr(AddressOf TimerProc))
        End If
    ElseIf lngTimerID <> 0 Then
        KillTimer Application.hWnd, lngTimerID
        lngTimerID = 0
    End If
    
    If isEnabled Xor isValue Then isEnabled = isValue
End Property

'// Raise event method
Friend Sub RaiseTimer()
    RaiseEvent Timer
End Sub
Private Sub Class_Terminate()
    KillTimer Application.hWnd, lngTimerID
End Sub

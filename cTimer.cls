VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class Timer"
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
Private blnEnabled_     As Boolean

'// Some properties
Friend Property Get Interval() As Long
    Interval = lngInterval
End Property
Friend Property Let Interval(ByVal lngValue As Long)
    If lngValue > 0 And lngValue <> lngInterval Then lngInterval = lngValue
End Property

Friend Property Get Enabled() As Boolean
    Enabled = blnEnabled_
End Property
Friend Property Let Enabled(ByVal blnEnabled As Boolean)
    If blnEnabled And lngInterval > 0 Then
        lngTimerID = SetTimer(Application.hWnd, ObjPtr(Me), lngInterval, GetPtr(AddressOf TimerProc))
    Else
        KillTimer Application.hWnd, lngTimerID
        lngTimerID = 0
    End If
    
    blnEnabled_ = blnEnabled
End Property

'// Raise event method
Friend Sub RaiseTimer()
    RaiseEvent Timer
End Sub
Private Sub Class_Terminate()
    KillTimer Application.hWnd, lngTimerID
End Sub

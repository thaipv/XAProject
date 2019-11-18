VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'       | |
'       | |
'   | | | |                           / /
'   |  || |__   ____  (_)  ____      / /
'   | | |  _ \ / _  | | | |  _ \ \  / /
'   | |_| | | | |_| |_| |_| |_| \ \/ /
'    \____| | |\__________   __/ \__/
'                         | |
'                         | |
'                         | |

'Author:    thaipv@live.com
'Date:      #19/09/2017#
'Note:      Flat scroller control
'Referent:  http://csdgn.org/article/scrollbar
'                                          _    .
'                                         |_|   BackControl
'                                               .
'      ____________________________        _    .
'     |                            |      | |   .
'     |                            |      | |   .
'     |                            |      | |   .
'     |            View            |      | |   ThumbControl
'     |          (Window)          |      | |   (Slider - Grip)
'     |                            |      | |   .
'     |                            |      | |   .
'     |____________________________|      |_|   .
'                                               .
'                                          _    .
'   _  ____________________________     _ |_|   NextControl
'  |_||____________________________|   |_|      .
'                                               .
'      minPoint  centerPoint  maxPoint          .
'        <------------*------------>            MovingSpace
'     <---------------*--------------->         ValueSpace
'   minValue     mediumValue      maxValue
'
'   f(x) = a(x-1) + b
'   thumbPos = ratio * (value - 1) + minPos

Public Event Change()

Private WithEvents fraMain_ As MSForms.Frame
Attribute fraMain_.VB_VarHelpID = -1
Private WithEvents lblBack_ As MSForms.Label
Attribute lblBack_.VB_VarHelpID = -1
Private WithEvents lblThum_ As MSForms.Label
Attribute lblThum_.VB_VarHelpID = -1
Private WithEvents lblNext_ As MSForms.Label
Attribute lblNext_.VB_VarHelpID = -1

Private minValue    As Integer
Private maxValue    As Integer
Private intValue    As Integer
Private intStep     As Integer

Private sngRatio    As Single
Private minSpace    As Single
Private maxSpace    As Single

'// Initialize
Private Sub Class_Initialize()
    minValue = 0: intStep = 1: maxValue = 1
End Sub

'// Let (set) and get all properties
Friend Property Let Min(ByRef ValueMin As Long)
    If minValue <> ValueMin Then
        minValue = ValueMin
        Call GetRatio
    End If
End Property
Friend Property Let Max(ByRef ValueMax As Long)
    If maxValue <> ValueMax Then
        maxValue = ValueMax
        Call GetRatio
    End If
End Property
Friend Property Let Step(ByRef ValueStep As Long)
    If intStep <> ValueStep Then intStep = ValueStep
End Property
Friend Property Let Value(ByVal Changed As Long)
    If intValue = Changed Then Exit Property
    Select Case Changed
        Case Is < minValue: intValue = minValue
        Case Is > maxValue: intValue = maxValue
        Case Else:          intValue = Changed
    End Select
    If Not fraMain_ Is Nothing Then Call OnScrolling
End Property
Friend Property Get Value() As Long: Value = intValue: End Property

'// Set casting for all controls
Public Sub Casting(ByRef fraMain As MSForms.Frame, ByRef lblBack As MSForms.Label, _
                   ByRef lblThum As MSForms.Label, ByRef lblNext As MSForms.Label)
    Set fraMain_ = fraMain: Set lblBack_ = lblBack
    Set lblThum_ = lblThum: Set lblNext_ = lblNext
    Call GetSpace: Call OnScrolling
End Sub

'// Scrolling center
Private Sub GetSpace()
    If fraMain_.Height > fraMain_.Width Then
        minSpace = lblBack_.Top + lblBack_.Height + lblThum_.Height / 2
        maxSpace = lblNext_.Top - lblThum_.Height / 2
    Else
        minSpace = lblBack_.Left + lblBack_.Width + lblThum_.Width / 2
        maxSpace = lblNext_.Left - lblThum_.Width / 2
    End If
    Call GetRatio
End Sub
Private Sub GetRatio()
    If maxValue = minValue Then
        sngRatio = maxSpace - minSpace
    Else
        sngRatio = (maxSpace - minSpace) / (maxValue - minValue)
    End If
End Sub
Private Sub OnScrolling()
    Dim midSpace As Single

    '// Get position value
    Select Case intValue
        Case minValue:  midSpace = minSpace
        Case maxValue:  midSpace = maxSpace
        Case Else
            midSpace = sngRatio * (intValue - 1) + minSpace
            midSpace = WorksheetFunction.Min(maxSpace, midSpace)
            midSpace = WorksheetFunction.Max(minSpace, midSpace)
    End Select

    '// Get position for #Thumbnail
    If fraMain_.Height > fraMain_.Width Then
        If lblThum_.Top <> midSpace - lblThum_.Height / 2 Then
            lblThum_.Top = midSpace - lblThum_.Height / 2
        End If
    Else
        If lblThum_.Left <> midSpace - lblThum_.Width / 2 Then
            lblThum_.Left = midSpace - lblThum_.Width / 2
        End If
    End If
    
    '// Raise change event
    RaiseEvent Change
End Sub
'// Handle events of some controls
Private Sub ClickCenter(Optional ByRef isPlus As Boolean)
    intValue = WorksheetFunction.Min(intValue, maxValue)
    intValue = WorksheetFunction.Max(intValue, minValue)
    Me.Value = intValue + IIf(isPlus, intStep, -intStep)
End Sub
'https://msdn.microsoft.com/VBA/Language-Reference-VBA/articles/mousedown-mouseup-events
Private Sub lblBack__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then Call ClickCenter(False) 'fmButtonLeft or fmButtonRight
End Sub
Private Sub lblNext__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then Call ClickCenter(True)
End Sub
Private Sub fraMain__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then
        Dim midSpace As Single, intChange As Integer

        midSpace = IIf(fraMain_.Height > fraMain_.Width, y - lblThum_.Height / 2, x - lblThum_.Width / 2)
        Select Case midSpace
            Case Is <= minSpace: intChange = minValue
            Case Is >= maxSpace: intChange = maxValue
            Case Else:  intChange = (midSpace - minSpace) / sngRatio + 1
        End Select
        Me.Value = intChange
    End If
End Sub
Private Sub lblThum__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then
        Dim midSpace As Single, intChange As Integer

        midSpace = IIf(fraMain_.Height > fraMain_.Width, lblThum_.Top + y, lblThum_.Left + x)
        intChange = (midSpace - minSpace) / sngRatio + 1
        
        Me.Value = intChange
    End If
End Sub
Private Sub fraMain__Layout()
    Dim posRequired As Single

    If fraMain_.Height > fraMain_.Width Then
        posRequired = (fraMain_.Height - lblNext_.Height) - lblBack_.Top
        If lblNext_.Top <> posRequired Then
            lblNext_.Top = posRequired: Call GetSpace: Call OnScrolling
        End If
    Else
        posRequired = (fraMain_.Width - lblNext_.Width) - lblBack_.Left
        If lblNext_.Left <> posRequired Then
            lblNext_.Left = posRequired: Call GetSpace: Call OnScrolling
        End If
    End If
End Sub

'// Terminate all variables
Private Sub Class_Terminate()
    If Not fraMain_ Is Nothing Then Set fraMain_ = Nothing
    If Not lblBack_ Is Nothing Then Set lblBack_ = Nothing
    If Not lblThum_ Is Nothing Then Set lblThum_ = Nothing
    If Not lblNext_ Is Nothing Then Set lblNext_ = Nothing
End Sub


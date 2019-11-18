VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Flat command button (include a frame, 2 label)"
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
'Date:      #12/12/2017#
'Note:      Flat command button (include a frame, 2 labels)

Public Event Click()

Private WithEvents cColor     As oColor
Attribute cColor.VB_VarHelpID = -1
Private WithEvents cTracker_  As eTracker
Attribute cTracker_.VB_VarHelpID = -1
Private WithEvents fraCover_  As MSForms.Frame
Attribute fraCover_.VB_VarHelpID = -1
Private WithEvents lblEmoji_  As MSForms.Label
Attribute lblEmoji_.VB_VarHelpID = -1
Private WithEvents lblTitle_  As MSForms.Label
Attribute lblTitle_.VB_VarHelpID = -1

Private isDefault_  As Boolean
Private isWarning_  As Boolean

'// Initialize
Private Sub Class_Initialize()
    Set cColor = oColor
End Sub

'// Get and let some properties
Friend Property Let isDefault(ByRef Value As Boolean):  isDefault_ = Value: End Property
Friend Property Let isWarning(ByRef Value As Boolean):  isWarning_ = Value: End Property
Friend Property Let Visible(ByRef Value As Boolean)
    If fraCover_ Is Nothing Then Exit Property
    If fraCover_.Visible Xor Value Then
        fraCover_.Visible = Value:  Call GetColor(False, False)
    End If
End Property
Friend Property Let Enabled(ByRef Value As Boolean)
    If fraCover_ Is Nothing Then Exit Property
    If fraCover_.Enabled Xor Value Then
        fraCover_.Enabled = Value:  Call GetColor(False, False)
    End If
End Property
Friend Property Get Enabled() As Boolean:   Enabled = fraCover_.Enabled:    End Property
Friend Property Get Visible() As Boolean:   Visible = fraCover_.Visible:    End Property

'// Set casting for all controls
Friend Sub Casting(fraCover As MSForms.Frame, lblEmoji As MSForms.Label, _
                   lblTitle As MSForms.Label, cTracker As eTracker)
    Set fraCover_ = fraCover:   Set lblEmoji_ = lblEmoji
    Set lblTitle_ = lblTitle:   Set cTracker_ = cTracker
    Call GetColor(False, False)
End Sub

'// Catch some events of objects
Private Sub cColor_ChangeColor()
    Call GetColor(False, False)
End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    Select Case fraCover_.Name
        Case strLeave:  Call GetColor(False, False)
        Case strHover:  Call GetColor(False, True)
    End Select
End Sub
Private Sub fraCover__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    cTracker_.OnEnter (fraCover_.Name)
End Sub
Private Sub lblEmoji__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    cTracker_.OnEnter (fraCover_.Name)
End Sub
Private Sub lblTitle__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    cTracker_.OnEnter (fraCover_.Name)
End Sub

'// Some methods
Private Sub GetColor(ByRef isClicked As Boolean, ByRef isHover As Boolean)
    Dim arrColors As Variant
    
    '// Get color for all controls
    With cColor
    Select Case True
        Case Not Me.Enabled: arrColors = Array(.clrFreeze, .clrBack, .clrFreeze)
        Case isClicked
            If isWarning_ Then
                arrColors = Array(.clrWarning, .clrWarning, .clrBack)
            Else
                arrColors = Array(.clrActive, .clrActive, .clrBack)
            End If
        Case isHover
            If isWarning_ Then
                arrColors = Array(.clrWarning, .clrWarning, .clrBack)
            Else
                arrColors = Array(.clrInstant, .clrInstant, .clrBack)
            End If
        Case isDefault_: arrColors = Array(.clrDefault, .clrDefault, .clrBack)
        Case Else:       arrColors = Array(.clrNormal, .clrBack, .clrNormal)
    End Select
    End With
    
    '// Fill color for all controls
    If fraCover_.BorderColor <> arrColors(0) Then
        fraCover_.BorderColor = arrColors(0): lblEmoji_.ForeColor = arrColors(2)
        fraCover_.BackColor = arrColors(1):   lblTitle_.ForeColor = arrColors(2)
    End If
End Sub
Private Sub OnClick()
    If Not Me.Enabled Then Exit Sub
    Call GetColor(True, True):  Call WaitingTime(126)
    Call GetColor(False, True): RaiseEvent Click
End Sub
Private Sub fraCover__Click():  Call OnClick:   End Sub
Private Sub lblEmoji__Click():  Call OnClick:   End Sub
Private Sub lblTitle__Click():  Call OnClick:   End Sub

'// Terminate
Private Sub Class_Terminate()
    Set cColor = Nothing
    If Not cTracker_ Is Nothing Then Set cTracker_ = Nothing
    If Not fraCover_ Is Nothing Then Set fraCover_ = Nothing
    If Not lblEmoji_ Is Nothing Then Set lblEmoji_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
End Sub

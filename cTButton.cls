VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Flat toggle button"
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
'Note:      Flat toggle button (Custom control)

Public Event Click()
Public Event RClick()

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

Private isPressed   As Boolean

'// Initialize
Private Sub Class_Initialize()
    Set cColor = oColor
End Sub

'// Get and let some properties
Friend Property Let Visible(ByRef Value As Boolean)
    If fraCover_ Is Nothing Then Exit Property
    If fraCover_.Visible Xor Value Then
        fraCover_.Visible = Value:  Call GetColor
    End If
End Property
Friend Property Let Enabled(ByRef Value As Boolean)
    If fraCover_ Is Nothing Then Exit Property
    If fraCover_.Enabled Xor Value Then
        fraCover_.Enabled = Value
        Call GetColor
    End If
End Property
Friend Property Let Pressed(ByRef Value As Boolean)
    If fraCover_ Is Nothing Then Exit Property
    If isPressed Xor Value Then
        isPressed = Value
        Call GetColor
    End If
End Property
Friend Property Get Visible() As Boolean:   Visible = fraCover_.Visible:    End Property
Friend Property Get Enabled() As Boolean:   Enabled = fraCover_.Enabled:    End Property
Friend Property Get Pressed() As Boolean:   Pressed = isPressed:            End Property

'// Set casting for all controls
Friend Sub Casting(fraCover As MSForms.Frame, lblEmoji As MSForms.Label, _
                   lblTitle As MSForms.Label, cTracker As eTracker)
    Set fraCover_ = fraCover:   Set lblEmoji_ = lblEmoji
    Set lblTitle_ = lblTitle:   Set cTracker_ = cTracker
    Call GetColor
End Sub

'// Catch some events of objects
Private Sub cColor_ChangeColor()
    Call GetColor
End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    Select Case fraCover_.Name
        Case strLeave:  Call GetColor(False)
        Case strHover:  Call GetColor(True)
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
Private Sub GetColor(Optional ByRef isHover As Boolean)
    Dim arrColors As Variant
     
    '// Get color for all controls
    With cColor
    Select Case True
        Case Not Me.Enabled: arrColors = Array(.clrFreeze, .clrBack, .clrFreeze)
        Case isPressed:      arrColors = Array(.clrActive, .clrActive, .clrBack)
        Case isHover:        arrColors = Array(.clrInstant, .clrInstant, .clrBack)
        Case Else:           arrColors = Array(.clrNormal, .clrBack, .clrNormal)
    End Select
    End With
    
    '// Fill color for all controls
    If fraCover_.BorderColor <> arrColors(0) Then
        fraCover_.BorderColor = arrColors(0): lblEmoji_.ForeColor = arrColors(2)
        fraCover_.BackColor = arrColors(1):   lblTitle_.ForeColor = arrColors(2)
    End If
End Sub

'// RaiseEvent LClick
Private Sub OnClick()
    If Not Me.Enabled Then Exit Sub

    isPressed = Not isPressed
    Call GetColor(isPressed)
    
    RaiseEvent Click
End Sub
Private Sub fraCover__Click():  Call OnClick:   End Sub
Private Sub lblEmoji__Click():  Call OnClick:   End Sub
Private Sub lblTitle__Click():  Call OnClick:   End Sub

'// RaiseEvent RClick
Private Sub fraCover__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = vbKeyRButton Then RaiseEvent RClick
End Sub
Private Sub lblEmoji__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = vbKeyRButton Then RaiseEvent RClick
End Sub
Private Sub lblTitle__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = vbKeyRButton Then RaiseEvent RClick
End Sub

'// Terminate
Private Sub Class_Terminate()
    Set cColor = Nothing
    If Not cTracker_ Is Nothing Then Set cTracker_ = Nothing
    If Not fraCover_ Is Nothing Then Set fraCover_ = Nothing
    If Not lblEmoji_ Is Nothing Then Set lblEmoji_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
End Sub

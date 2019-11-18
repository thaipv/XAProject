VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFlatTab"
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
'Note:      Flat tab (include a frame & 2 labels)

Public Event Activate()
Public Event DeActivate()

Private WithEvents cColor     As oColor
Attribute cColor.VB_VarHelpID = -1
Private WithEvents cTracker_  As eTracker
Attribute cTracker_.VB_VarHelpID = -1
Private WithEvents cCentral_  As eAction
Attribute cCentral_.VB_VarHelpID = -1

Private WithEvents fraCover_  As MSForms.Frame
Attribute fraCover_.VB_VarHelpID = -1
Private WithEvents lblEmoji_  As MSForms.Label
Attribute lblEmoji_.VB_VarHelpID = -1
Private WithEvents lblTitle_  As MSForms.Label
Attribute lblTitle_.VB_VarHelpID = -1

'// Initialize
Private Sub Class_Initialize()
    Set cColor = oColor
End Sub

'// Get and let some properties
Friend Property Let Visible(ByRef Value As Boolean)
    If fraCover_ Is Nothing Then Exit Property
    If fraCover_.Visible Xor Value Then
        fraCover_.Visible = Value
        If fraCover_.Visible Then Call GetColor
    End If
End Property
Friend Property Get Visible() As Boolean:   Visible = fraCover_.Visible:    End Property

'// Set casting for all controls
Friend Sub Casting(fraCover As MSForms.Frame, lblEmoji As MSForms.Label, lblTitle As MSForms.Label, _
                   cTracker As eTracker, cCentral As eAction)
    Set fraCover_ = fraCover:   Set lblEmoji_ = lblEmoji:   Set lblTitle_ = lblTitle
    Set cTracker_ = cTracker:   Set cCentral_ = cCentral
    Call GetColor
End Sub

'// Catch some events of objects
Private Sub cColor_ChangeColor()
    Call GetColor
End Sub
Private Sub cCentral__OnChange(ByRef strDeActive As String, ByRef strActive As String)
    Select Case fraCover_.Name
        Case strDeActive: Call ActionCenter(False, False)
        Case strActive:   Call ActionCenter(True, False)
    End Select
End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    Select Case fraCover_.Name
        Case strLeave:  Call GetColor(False)
        Case strHover:  Call GetColor(True)
    End Select
End Sub
Private Sub fraCover__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    Call cTracker_.OnEnter(fraCover_.Name)
End Sub
Private Sub lblEmoji__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    Call cTracker_.OnEnter(fraCover_.Name)
End Sub
Private Sub lblTitle__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    Call cTracker_.OnEnter(fraCover_.Name)
End Sub

'// Some methods
Private Sub GetColor(Optional ByRef isHover As Boolean, Optional isActive As Boolean)
    If Not Me.Visible Then Exit Sub
    
    Dim clrFont As Long
    '// Get color for all controls
    If isHover Then
        clrFont = cColor.clrInstant
    ElseIf isActive Then
        clrFont = cColor.clrActive
    Else
        clrFont = cColor.clrNormal
    End If

    '// Fill color for all controls
    If lblEmoji_.ForeColor <> clrFont Then
        lblEmoji_.ForeColor = clrFont: lblTitle_.ForeColor = clrFont
    End If
End Sub
Private Sub ActionCenter(ByRef isActive As Boolean, ByRef isClicked As Boolean)
    If isActive Then RaiseEvent Activate Else RaiseEvent DeActivate
    Call GetColor(isClicked, isActive)
End Sub
Friend Sub Activate():  Call cCentral_.OnAction(fraCover_.Name): End Sub
Private Sub fraCover__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then Call Activate
End Sub
Private Sub lblEmoji__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then Call Activate
End Sub
Private Sub lblTitle__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then Call Activate
End Sub

'// Terminate
Private Sub Class_Terminate()
    Set cColor = Nothing
    If Not cTracker_ Is Nothing Then Set cTracker_ = Nothing
    If Not cCentral_ Is Nothing Then Set cCentral_ = Nothing
    If Not fraCover_ Is Nothing Then Set fraCover_ = Nothing
    If Not lblEmoji_ Is Nothing Then Set lblEmoji_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
End Sub

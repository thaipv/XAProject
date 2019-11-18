VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEditBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Flat textbox"
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
'Note:      Flat textbox (Custom control)

Public Event Change()

Private WithEvents cColor       As oColor
Attribute cColor.VB_VarHelpID = -1
Private WithEvents cTracker_    As eTracker
Attribute cTracker_.VB_VarHelpID = -1
Private WithEvents fraCover_    As MSForms.Frame
Attribute fraCover_.VB_VarHelpID = -1
Private WithEvents txtInput_    As MSForms.TextBox
Attribute txtInput_.VB_VarHelpID = -1
Private lblTitle_               As MSForms.Label

'// Initialize
Private Sub Class_Initialize()
    Set cColor = oColor
End Sub

'// Get and let some properties
Friend Property Let Enabled(ByRef Value As Boolean)
    If fraCover_.Enabled Xor Value Then
        fraCover_.Enabled = Value
        Call GetColor(False)
    End If
End Property
Friend Property Let Value(ByRef strText As String)
    If fraCover_ Is Nothing Then Exit Property
    If txtInput_.Value <> strText Then txtInput_.Value = strText
End Property
Friend Property Let Locked(ByRef Value As Boolean): txtInput_.Locked = Value:       End Property
Friend Property Get Enabled() As Boolean:           Enabled = fraCover_.Enabled:    End Property
Friend Property Get Locked() As Boolean:            Locked = txtInput_.Locked:      End Property
Friend Property Get Value() As String:              Value = txtInput_.Value:        End Property

'// Set casting for all controls
Friend Sub Casting(fraCover As MSForms.Frame, txtInput As MSForms.TextBox, _
                   lblNotice As MSForms.Label, cTracker As eTracker)
    Set fraCover_ = fraCover:   Set txtInput_ = txtInput
    Set lblTitle_ = lblNotice:  Set cTracker_ = cTracker
    Call GetColor(False)
End Sub

'// Catch some events of objects
Private Sub cColor_ChangeColor()
    Call GetColor(False)
End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    If fraCover_.Name = strLeave Then
        Call GetColor(False)
    ElseIf fraCover_.Name = strHover Then
        Call GetColor(True)
    End If
End Sub
Private Sub fraCover__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    Call cTracker_.OnEnter(fraCover_.Name)
End Sub
Private Sub txtInput__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    Call cTracker_.OnEnter(fraCover_.Name)
End Sub
Private Sub fraCover__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    txtInput_.SelStart = 0
    txtInput_.SelLength = txtInput_.TextLength
End Sub
Private Sub txtInput__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call fraCover__DblClick(Cancel)
End Sub
Private Sub txtInput__Change()
    Dim isShowNote As Boolean

    Call cTracker_.OnEnter(fraCover_.Name)
    isShowNote = (Len(txtInput_.Value) = 0)
    If lblTitle_.Visible Xor isShowNote Then lblTitle_.Visible = isShowNote
    'strValue = txtInput_.Text

    RaiseEvent Change
End Sub

'// Some methods (SetFocus and GetColor)
Private Sub GetColor(ByRef isHover As Boolean)
    Dim lngColor As Long

    Select Case True
        Case Not Me.Enabled: lngColor = cColor.clrFreeze
        Case isHover:        lngColor = cColor.clrActive
        Case Else:           lngColor = cColor.clrNormal
    End Select
    
    If fraCover_.BorderColor <> lngColor Then
        fraCover_.BorderColor = lngColor:   txtInput_.ForeColor = lngColor
        lblTitle_.ForeColor = lngColor
    End If
End Sub

'// Terminate
Private Sub Class_Terminate()
    Set cColor = Nothing
    If Not cTracker_ Is Nothing Then Set cTracker_ = Nothing
    If Not fraCover_ Is Nothing Then Set fraCover_ = Nothing
    If Not txtInput_ Is Nothing Then Set txtInput_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
End Sub

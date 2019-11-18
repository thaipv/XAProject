VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOpenBox"
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

Private WithEvents cTracker_    As eTracker
Attribute cTracker_.VB_VarHelpID = -1
Private lblTitle_               As MSForms.Label
Attribute lblTitle_.VB_VarHelpID = -1
Private WithEvents txtInput_    As MSForms.TextBox
Attribute txtInput_.VB_VarHelpID = -1
Private lblULine_               As MSForms.Label
Attribute lblULine_.VB_VarHelpID = -1

'// Get and let some properties
Friend Property Let Value(ByRef strText As String)
    If txtInput_ Is Nothing Then Exit Property
    If txtInput_.Value <> strText Then txtInput_.Value = strText
End Property
Friend Property Let Locked(ByRef Value As Boolean)
    If txtInput_ Is Nothing Then Exit Property
    If txtInput_.Locked Xor Value Then txtInput_.Locked = Value
End Property
Friend Property Get Locked() As Boolean:    Locked = txtInput_.Locked:  End Property
Friend Property Get Value() As String:      Value = txtInput_.Value:    End Property

'// Set casting for all controls
Friend Sub Casting(ByRef lblTitle As MSForms.Label, ByRef txtInput As MSForms.TextBox, _
                   ByRef lblUline As MSForms.Label, ByRef cTracker As eTracker)
    Set lblTitle_ = lblTitle:   Set txtInput_ = txtInput
    Set lblULine_ = lblUline:   Set cTracker_ = cTracker
    Call GetColor(False)
End Sub

'// Catch some events of objects
'Private Sub cColor_ChangeColor()
'    Call GetColor(False)
'End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    If lblTitle_.Name = strLeave Then
        Call GetColor(False)
    ElseIf lblTitle_.Name = strHover Then
        Call GetColor(True)
    End If
End Sub
Private Sub txtInput__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    txtInput_.SelStart = 0
    txtInput_.SelLength = txtInput_.TextLength
End Sub
Private Sub txtInput__Change()
    Call cTracker_.OnEnter(lblTitle_.Name)
    RaiseEvent Change
End Sub

'// Some methods (SetFocus and GetColor)
Private Sub GetColor(ByRef isHover As Boolean)
    Dim clrFont As Long

    clrFont = IIf(isHover, oColor.clrActive, oColor.clrNormal)
    If lblTitle_.ForeColor <> clrFont Then
        lblTitle_.ForeColor = clrFont: txtInput_.ForeColor = clrFont
        lblULine_.BackColor = clrFont
    End If
End Sub

'// Terminate
Private Sub Class_Terminate()
    If Not cTracker_ Is Nothing Then Set cTracker_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
    If Not txtInput_ Is Nothing Then Set txtInput_ = Nothing
    If Not lblULine_ Is Nothing Then Set lblULine_ = Nothing
End Sub

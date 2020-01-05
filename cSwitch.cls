VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Flat switch (Custom control)"
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
'Note:      Flat switch (Custom control)

Public Event Change()

Private WithEvents cColor     As oColor
Attribute cColor.VB_VarHelpID = -1
Private WithEvents lblBorder_ As MSForms.Label
Attribute lblBorder_.VB_VarHelpID = -1
Private WithEvents lblFilled_ As MSForms.Label
Attribute lblFilled_.VB_VarHelpID = -1
Private WithEvents lblSlider_ As MSForms.Label
Attribute lblSlider_.VB_VarHelpID = -1
Private OnOff                 As Boolean

'// Initialize
Private Sub Class_Initialize()
    Set cColor = oColor
End Sub

'// Get and let properties
Friend Property Let Status(ByRef Value As Boolean)
    If OnOff Xor Value Then
        OnOff = Value: Call OnAction(True)
    Else
        Call OnAction(False)
    End If
End Property
Friend Property Let Enabled(ByRef Value As Boolean)
    If lblBorder_.Enabled Xor Value Then
        lblBorder_.Enabled = Value: lblFilled_.Enabled = Value
        lblSlider_.Enabled = Value: Call GetColor
    End If
End Property
Friend Property Get Status() As Boolean
    Status = OnOff
End Property
Friend Property Get Enabled() As Boolean
    Enabled = lblBorder_.Enabled
End Property

'// Set casting for all controls
Friend Sub Casting(lblBorder As MSForms.Label, lblFilled As MSForms.Label, lblSlider As MSForms.Label)
    Set lblBorder_ = lblBorder: Set lblFilled_ = lblFilled
    Set lblSlider_ = lblSlider: Call OnAction(False)
End Sub

'// Catch some events of objects
Private Sub lblBorder__Click():
    If Me.Enabled Then Me.Status = Not Me.Status
End Sub
Private Sub lblFilled__Click(): Call lblBorder__Click: End Sub
Private Sub lblSlider__Click(): Call lblBorder__Click: End Sub
Private Sub cColor_ChangeColor(): Call GetColor:       End Sub

'// Action center
Private Sub OnAction(Optional ByRef isReset As Boolean = False)
    Dim sngLeft As Single  'Property Left of lblSlider_
    
    sngLeft = lblBorder_.Left + IIf(Me.Status, lblBorder_.Width - lblSlider_.Width, 0)
    If lblSlider_.Left <> sngLeft Then lblSlider_.Left = sngLeft
    If lblFilled_.Visible Xor Me.Status Then lblFilled_.Visible = Me.Status
    
    Call GetColor
    If isReset Then RaiseEvent Change
End Sub

'// Color center
Private Sub GetColor()
    Dim arrColors As Variant
    
    '// Get color for all controls
    If Not Me.Enabled Then
        arrColors = Array(cColor.clrFreeze, cColor.clrFreeze, cColor.clrBack)
    Else
        arrColors = Array(cColor.clrNormal, cColor.clrActive, cColor.clrBack)
    End If
    
    '// Fill color for all controls
    If lblBorder_.BorderColor <> arrColors(0) Then
        lblBorder_.BorderColor = arrColors(0):  lblFilled_.BackColor = arrColors(1)
        lblSlider_.BackColor = arrColors(0):    lblSlider_.BorderColor = arrColors(2)
    End If
End Sub

'// Terminate
Private Sub Class_Terminate()
    Set cColor = Nothing
    If Not lblBorder_ Is Nothing Then Set lblBorder_ = Nothing
    If Not lblFilled_ Is Nothing Then Set lblFilled_ = Nothing
    If Not lblSlider_ Is Nothing Then Set lblSlider_ = Nothing
End Sub

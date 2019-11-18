VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cListBox"
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
'Date:      #12/12/2017#
'Note:      Flat textbox (Custom control)

Public Event DropDown(ByRef isDropList As Boolean)

Private WithEvents cColor       As oColor
Attribute cColor.VB_VarHelpID = -1
Private WithEvents cTracker_    As eTracker
Attribute cTracker_.VB_VarHelpID = -1
Private WithEvents fraCover_    As MSForms.Frame
Attribute fraCover_.VB_VarHelpID = -1
Private WithEvents txtInput_    As MSForms.TextBox
Attribute txtInput_.VB_VarHelpID = -1
Private WithEvents lblDropX_    As MSForms.Label
Attribute lblDropX_.VB_VarHelpID = -1
Private lblTitle_               As MSForms.Label

Private isDropList  As Boolean
Private strStyle    As String
Private strValue    As String

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
Friend Property Let Format(ByRef strValue As String)
    strStyle = strValue
End Property
Friend Property Let Value(ByRef strText As String)
    strValue = strText
    txtInput_.Text = IIf(strStyle = vbNullString, strValue, Strings.Format$(strValue, strStyle))
    Call GetColor(True)
End Property
Friend Property Get Enabled() As Boolean:   Enabled = fraCover_.Enabled:    End Property
Friend Property Get Value() As String:      Value = strValue:               End Property

'// Set casting for all controls
Friend Sub Casting(ByRef fraCover As MSForms.Frame, ByRef txtInput As MSForms.TextBox, _
                   ByRef lblNotice As MSForms.Label, ByRef lblDropButton As MSForms.Label, _
                   ByRef cTracker As eTracker)
    Set fraCover_ = fraCover:       Set txtInput_ = txtInput:   Set lblTitle_ = lblNotice
    Set lblDropX_ = lblDropButton:  Set cTracker_ = cTracker:   Call GetColor(False)
End Sub

'// Catch some events of objects
Private Sub cColor_ChangeColor()
    Call GetColor(False)
End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    If fraCover_.Name = strLeave Then
        Call GetDropList(False)
        Call GetVisible(txtInput_.Text = vbNullString, False)
        Call GetColor(False)
    ElseIf fraCover_.Name = strHover Then
        Call GetVisible(txtInput_.Text = vbNullString, True)
        Call GetColor(True)
    End If
End Sub
Private Sub fraCover__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    Call cTracker_.OnEnter(fraCover_.Name)
End Sub
Private Sub lblDropX__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then Call GetDropList(Not isDropList)
    Call GetColor(True)
End Sub
Private Sub txtInput__Change()
    Call GetVisible(Len(txtInput_.Text) = 0, isDropList)
    'strValue = txtInput_.Text '// Using Let Property only
End Sub

'// Some methods (SetFocus and GetColor)
Private Sub GetDropList(ByRef isDropDown As Boolean)
    If isDropList Xor isDropDown Then
        isDropList = isDropDown
        RaiseEvent DropDown(isDropList)
    End If
End Sub
Private Sub GetColor(ByRef isHover As Boolean)
    Dim lngColor As Long

    '// Get color for fraCover, txtInput, lblTitle
    Select Case True
        Case Not Me.Enabled: lngColor = cColor.clrFreeze
        Case isHover:        lngColor = cColor.clrActive
        Case Else:           lngColor = cColor.clrNormal
    End Select

    If fraCover_.BorderColor <> lngColor Then
        fraCover_.BorderColor = lngColor: txtInput_.ForeColor = lngColor
        lblTitle_.ForeColor = lngColor
    End If
    
    '// Get color for lblDropButton
    If lblDropX_.Visible Then
        lngColor = IIf(isDropList, cColor.clrActive, cColor.clrNormal)
        If lblDropX_.ForeColor <> lngColor Then lblDropX_.ForeColor = lngColor
    End If
End Sub
Private Sub GetVisible(ByRef isShowNote As Boolean, ByRef isDropList As Boolean)
    If lblTitle_.Visible Xor isShowNote Then lblTitle_.Visible = isShowNote
    If lblDropX_.Visible Xor isDropList Then lblDropX_.Visible = isDropList
End Sub

'// Terminate
Private Sub Class_Terminate()
    Set cColor = Nothing
    If Not cTracker_ Is Nothing Then Set cTracker_ = Nothing
    If Not fraCover_ Is Nothing Then Set fraCover_ = Nothing
    If Not txtInput_ Is Nothing Then Set txtInput_ = Nothing
    If Not lblDropX_ Is Nothing Then Set lblDropX_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
End Sub

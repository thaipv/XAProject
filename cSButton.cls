VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Flat simple button (include a label)"
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
'Note:      Flat simple button (include a label)

Public Event Click()

Private WithEvents cColor       As oColor
Attribute cColor.VB_VarHelpID = -1
Private WithEvents cTracker_    As eTracker
Attribute cTracker_.VB_VarHelpID = -1
Private WithEvents lblTitle_    As MSForms.Label
Attribute lblTitle_.VB_VarHelpID = -1
Private isWarning_              As Boolean
Private lngNatural              As Long

'// Initialize
Private Sub Class_Initialize()
    Set cColor = oColor
    lngNatural = cColor.clrNormal
End Sub

'// Get and let some properties
Friend Property Let isWarning(ByRef Value As Boolean)
    If isWarning_ Xor Value Then isWarning_ = Value
End Property
Friend Property Let NaturalColor(ByRef Value As Long)
    If lngNatural <> Value Then lngNatural = Value
    If Not lblTitle_ Is Nothing Then Call GetColor(False, False)
End Property
Friend Property Let Visible(ByRef Value As Boolean)
    If lblTitle_ Is Nothing Then Exit Property
    If lblTitle_.Visible Xor Value Then
        lblTitle_.Visible = Value
        Call GetColor(False, False)
    End If
End Property
Friend Property Let Enabled(ByRef Value As Boolean)
    If lblTitle_ Is Nothing Then Exit Property
    If lblTitle_.Enabled Xor Value Then
        lblTitle_.Enabled = Value
        Call GetColor(False, False)
    End If
End Property
Friend Property Get Enabled() As Boolean:   Enabled = lblTitle_.Enabled:    End Property
Friend Property Get Visible() As Boolean:   Visible = lblTitle_.Visible:    End Property

'// Set casting for all controls
Friend Sub Casting(ByRef lblTitle As MSForms.Label, ByRef cTracker As eTracker)
    Set lblTitle_ = lblTitle
    Set cTracker_ = cTracker
    Call GetColor(False, False)
End Sub

'// Catch some events of objects
Private Sub cColor_ChangeColor()
    Call GetColor(False, False)
End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    Select Case lblTitle_.Name
        Case strLeave:  Call GetColor(False, False)
        Case strHover:  Call GetColor(False, True)
    End Select
End Sub
Private Sub lblTitle__MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    cTracker_.OnEnter (lblTitle_.Name)
End Sub

'// Some methods
Private Sub GetColor(ByRef isClicked As Boolean, ByRef isHover As Boolean)
    Dim clrFont As Long
    
    '// Get color for title control
    With cColor
    Select Case True
        Case Not Me.Enabled: clrFont = .clrFreeze
        Case isClicked:      clrFont = IIf(isWarning_, .clrWarning, .clrActive)
        Case isHover:        clrFont = IIf(isWarning_, .clrWarning, .clrInstant)
        Case Else:           clrFont = lngNatural
    End Select
    End With
    
    '// Fill color for title control
    If lblTitle_.ForeColor <> clrFont Then lblTitle_.ForeColor = clrFont
End Sub
Private Sub lblTitle__Click()
    If Not Me.Enabled Then Exit Sub
    Call GetColor(True, True):  'Call WaitingTime(126)
    Call GetColor(False, True): RaiseEvent Click
End Sub

'// Terminate
Private Sub Class_Terminate()
    Set cColor = Nothing
    If Not cTracker_ Is Nothing Then Set cTracker_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
End Sub

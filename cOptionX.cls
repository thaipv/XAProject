VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOptionX"
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
'Note:      Flat Option button (include 2 labels)

Public Event Change(ByRef isSelected As Boolean)

Private WithEvents cCentral_  As eAction
Attribute cCentral_.VB_VarHelpID = -1
Private WithEvents lblEmoji_  As MSForms.Label
Attribute lblEmoji_.VB_VarHelpID = -1
Private WithEvents lblTitle_  As MSForms.Label
Attribute lblTitle_.VB_VarHelpID = -1

'// Get and let some properties
Friend Property Let Value(ByRef isSelected As Boolean)
    If lblTitle_ Is Nothing Then Exit Property
    If isSelected Xor Me.Value Then
        If isSelected Then
            Call cCentral_.OnAction(lblTitle_.Name)
        Else
            Call OnChange(False, False)
        End If
    End If
End Property
Friend Property Get Value() As Boolean
    If lblTitle_ Is Nothing Then Exit Property
    Value = (cCentral_.ActiveControl = lblTitle_.Name)
End Property

'// Set casting for all controls
Friend Sub Casting(ByRef lblEmoji As MSForms.Label, ByRef lblTitle As MSForms.Label, _
                   ByRef cCentral As eAction)
    Set lblEmoji_ = lblEmoji:   Set lblTitle_ = lblTitle
    Set cCentral_ = cCentral:   Call GetColor(False)
End Sub

'// Catch some events of objects
Private Sub cCentral__OnChange(ByRef strDeActive As String, ByRef strActive As String)
    Select Case lblTitle_.Name
        Case strDeActive: Call OnChange(False, False)
        Case strActive:   Call OnChange(True, True)
    End Select
End Sub

'// Some methods
Private Sub GetColor(ByRef isHighlight As Boolean)
    Dim clrFont As Long
    
    clrFont = IIf(isHighlight, oColor.clrActive, oColor.clrNormal)
    If lblEmoji_.ForeColor <> clrFont Then
        lblEmoji_.ForeColor = clrFont: lblTitle_.ForeColor = clrFont
    End If
End Sub
Private Sub OnChange(ByRef isSelected As Boolean, ByRef isHighlight As Boolean)
    Dim strEmoji As String
    
    '// Change emoji for OptionButton
    strEmoji = IIf(isSelected, ChrW$(60619), ChrW$(60618))
    If lblEmoji_.Caption <> strEmoji Then lblEmoji_.Caption = strEmoji
    
    '// Create event & Highlight controls
    RaiseEvent Change(isSelected)
    If isHighlight And isSelected Then
        Call GetColor(True):    Call WaitingTime(210)
        Call GetColor(False)
    End If
End Sub
Private Sub lblEmoji__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then Call cCentral_.OnAction(lblTitle_.Name)
End Sub
Private Sub lblTitle__MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If Button = 1 Or Button = 2 Then Call cCentral_.OnAction(lblTitle_.Name)
End Sub

'// Terminate
Private Sub Class_Terminate()
    If Not cCentral_ Is Nothing Then Set cCentral_ = Nothing
    If Not lblEmoji_ Is Nothing Then Set lblEmoji_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
End Sub

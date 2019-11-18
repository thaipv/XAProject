VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDateBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Author:    HuuThang_BD (giaiphapexcel.com)
'Note:      Datebox with input mask (include 1 frame, 1 textbox and 2 labels)
'Source:    https://www.giaiphapexcel.com/diendan/threads/130366
'Edit by:   @thaipv #15/03/2018#

Public Event Change()
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
Private strStyle    As String 'dd/mm/yyyy
Private datValue    As Long

'// Initialize
Private Sub Class_Initialize()
    Set cColor = oColor
    strStyle = cSystem.FormatDate
End Sub

'// Get and let some properties
Friend Property Let Enabled(ByRef Value As Boolean)
    If fraCover_.Enabled Xor Value Then
        fraCover_.Enabled = Value
        Call GetColor(False)
    End If
End Property
Friend Property Let Value(ByRef datChanged As Long)
    If datValue <> datChanged Then
        datValue = datChanged
        txtInput_.Text = Format$(datValue, strStyle)
        Call GetColor(True)
    End If
End Property
Friend Property Let Locked(Value As Boolean):   txtInput_.Locked = Value:       End Property
Friend Property Get Locked() As Boolean:        Locked = txtInput_.Locked:      End Property
Friend Property Get Value() As Long:            Value = datValue:               End Property
Friend Property Get Enabled() As Boolean:       Enabled = fraCover_.Enabled:    End Property

'// Set casting for all controls
Friend Sub Casting(ByRef fraCover As MSForms.Frame, ByRef txtInput As MSForms.TextBox, _
                   ByRef lblNotice As MSForms.Label, ByRef lblDropButton As MSForms.Label, _
                   ByRef cTracker As eTracker)
    Set fraCover_ = fraCover:       Set txtInput_ = txtInput:   Set lblTitle_ = lblNotice
    Set lblDropX_ = lblDropButton:  Set cTracker_ = cTracker
    
    'txtInput_.Value = strStyle
    txtInput_.AutoTab = False
    txtInput_.MaxLength = Len(strStyle)
    txtInput_.SelStart = 0
    
    Call GetColor(False)
End Sub

'// Catch some events of objects
Private Sub cColor_ChangeColor()
    Call GetColor(False)
End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    If fraCover_.Name = strLeave Then
        txtInput_.Text = Format$(datValue, strStyle)
        Call GetDropList(False)
        Call GetVisible(txtInput_.TextLength = 0, False)
        Call GetColor(False)
    ElseIf fraCover_.Name = strHover Then
        Call GetVisible(txtInput_.TextLength = 0, True)
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
    Call cTracker_.OnEnter(fraCover_.Name)
End Sub
Private Sub txtInput__KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim strKeyChar As String, strValue As String, strTemp As String, bytCursor As Byte
    
    Select Case True
    Case KeyCode = vbKeyTab     'System default
    Case KeyCode = vbKeyUp:     KeyCode = vbKeyLeft
    Case KeyCode = vbKeyDown:   KeyCode = vbKeyRight
    Case KeyCode = vbKeyLeft    'System default
    Case KeyCode = vbKeyRight   'System default
    Case KeyCode = vbKeyHome    'System default
    Case KeyCode = vbKeyEnd     'System default
    Case Shift = 2 And KeyCode = 65 'Ctrl+A
    Case InStr(1, "0123456789", Chr$(KeyCode)) > 0
        strKeyChar = Chr(KeyCode)
        bytCursor = txtInput_.SelStart
        If txtInput_.TextLength = 0 Then txtInput_.Value = strStyle
        strValue = txtInput_.Value
        
        Select Case bytCursor
            Case 0 '_d/mm/yyyy
                If strKeyChar > 3 Then
                    Mid$(strValue, bytCursor + 1, 2) = "0" & strKeyChar
                    bytCursor = bytCursor + 3
                Else
                    Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                    bytCursor = bytCursor + 1
                End If
            Case 1 'd_/mm/yyyy
                Select Case Mid(strValue, bytCursor, 1)
                    Case 0, 1, 2
                        Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                        bytCursor = bytCursor + 2
                    Case 3
                        If strKeyChar < 2 Then
                            Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                            bytCursor = bytCursor + 2
                        Else
                            KeyCode = 0
                        End If
                End Select
            Case 2, 5 '[dd_mm/yyyy],[dd/mm_yyyy]
                bytCursor = bytCursor + 1
            Case 3 'dd/_m/yyyy
                If strKeyChar > 1 Then
                    Mid$(strValue, bytCursor + 1, 2) = "0" & strKeyChar
                    bytCursor = bytCursor + 3
                Else
                    Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                    bytCursor = bytCursor + 1
                End If
            Case 4 'dd/m_/yyyy
                Select Case Mid(strValue, bytCursor, 1)
                    Case 0
                        If strKeyChar > 0 Then
                            Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                            bytCursor = bytCursor + 2
                        Else
                            KeyCode = 0
                        End If
                    Case 1
                        If strKeyChar < 3 Then
                            Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                            bytCursor = bytCursor + 2
                        Else
                            KeyCode = 0
                        End If
                End Select
            Case 6 'dd/mm_yyyy
                If strKeyChar = 1 Or strKeyChar = 2 Then
                    Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                    bytCursor = bytCursor + 1
                Else
                    KeyCode = 0
                End If
            Case 7 'dd/mm/y_yy
                If Mid(strValue, bytCursor, 1) = 1 Then
                    If strKeyChar = 9 Then
                        Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                        bytCursor = bytCursor + 1
                    Else
                        KeyCode = 0
                    End If
                ElseIf Mid(strValue, bytCursor, 1) = 2 Then
                    If strKeyChar = 0 Then
                        Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                        bytCursor = bytCursor + 1
                    Else
                        KeyCode = 0
                    End If
                End If
            Case 8, 9 '[dd/mm/yy_y],[dd/mm/yyy_]
                Mid$(strValue, bytCursor + 1, 1) = strKeyChar
                bytCursor = bytCursor + 1
        End Select
        
        '// Let auto tab for txtInput and change value
        strTemp = Replace$(strValue, "/", "")
        If IsNumeric(strTemp) Then Call OnChange(strValue)
        txtInput_.AutoTab = (IsNumeric(strTemp) And bytCursor > 9)
        
        '// Let value and put cursor fot txtInput_
        txtInput_.Value = strValue
        txtInput_.SelStart = bytCursor
        KeyCode = 0
    Case KeyCode = vbKeyBack, KeyCode = vbKeyDelete 'BACKSPACE, DELETE
        Dim strSelText As String
        
        strValue = txtInput_.Value
        strSelText = txtInput_.SelText
        bytCursor = txtInput_.SelStart
        
        If bytCursor > 0 Then
            If Len(strSelText) = 0 Then
                Mid$(strValue, bytCursor, 1) = Mid$(strStyle, bytCursor, 1)
            Else
                Mid$(strValue, bytCursor + 1, Len(strSelText)) = _
                Mid$(strStyle, bytCursor + 1, Len(strSelText))
                bytCursor = bytCursor + 1
            End If
        Else
            If Len(strSelText) > 0 Then
                Mid$(strValue, bytCursor + 1, Len(strSelText)) = _
                Mid$(strStyle, bytCursor + 1, Len(strSelText))
            End If
        End If
        
        '// Let auto tab and value to txtInput_
        strTemp = Replace$(strValue, "/", "")
        If IsNumeric(strTemp) Then Call OnChange(strValue)
        txtInput_.AutoTab = (IsNumeric(strTemp) And bytCursor - 1 > 9)
        
        '// Let value and put cursor fot txtInput_
        txtInput_.Value = strValue
        txtInput_.SelStart = WorksheetFunction.Max(bytCursor - 1, 0)
        KeyCode = 0
    Case Else
        KeyCode = 0 'Do nothing
    End Select
End Sub

'// Some methods
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
    
    '// Fill color for all controls
    If fraCover_.BorderColor <> lngColor Then
        fraCover_.BorderColor = lngColor:   txtInput_.ForeColor = lngColor
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
Private Sub OnChange(ByRef strValue As String)
    datValue = DateSerial(Val(Mid(strValue, 7, 4)), Val(Mid(strValue, 4, 2)), _
                          Val(Mid(strValue, 1, 2)))
    RaiseEvent Change
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

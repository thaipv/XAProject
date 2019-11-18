VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDateRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Author:    HuuThang_BD (giaiphapexcel.com)
'Note:      Open datebox with input mask
'Source:    https://www.giaiphapexcel.com/diendan/threads/130366
'Edit by:   @thaipv #15/03/2018#

Public Event Change()

Private WithEvents cTracker_    As eTracker
Attribute cTracker_.VB_VarHelpID = -1
Private lblTitle_               As MSForms.Label
Private lblULine_               As MSForms.Label
Private WithEvents txtInput_    As MSForms.TextBox
Attribute txtInput_.VB_VarHelpID = -1

Private strStyle                As String 'dd/mm/yyyy
Private datValue                As Long

'// Initialize
Private Sub Class_Initialize()
    strStyle = cSystem.FormatDate
End Sub

'// Get and let some properties
Friend Property Let Value(ByRef datChanged As Long)
    If datValue <> datChanged Then
        datValue = datChanged
        txtInput_.Text = Format$(datValue, strStyle)
        Call OnChange(Format$(datValue, strStyle))
        Call GetColor(True)
    End If
End Property
Friend Property Let Locked(ByRef Value As Boolean): txtInput_.Locked = Value:   End Property
Friend Property Get Locked() As Boolean:            Locked = txtInput_.Locked:  End Property
Friend Property Get Value() As Long:                Value = datValue:           End Property

'// Set casting for all controls
Public Sub Casting(ByRef lblTitle As MSForms.Label, ByRef txtInput As MSForms.TextBox, _
                   ByRef lblUline As MSForms.Label, ByRef cTracker As eTracker)
    Set lblTitle_ = lblTitle:   Set txtInput_ = txtInput
    Set lblULine_ = lblUline:   Set cTracker_ = cTracker
    
    txtInput_.Value = strStyle
    txtInput_.AutoTab = False
    txtInput_.MaxLength = Len(strStyle)
    txtInput_.SelStart = 0
    
    Call GetColor(False)
End Sub

'// Catch some events of objects
'Private Sub cColor_ChangeColor()
'    Call GetColor(False)
'End Sub
Private Sub cTracker__MouseTrack(ByRef strLeave As String, ByRef strHover As String)
    If lblTitle_.Name = strLeave Then
        txtInput_.Text = Format$(datValue, strStyle)
        Call GetColor(False)
    ElseIf lblTitle_.Name = strHover Then
        Call GetColor(True)
    End If
End Sub
Private Sub txtInput__Change()
    Call cTracker_.OnEnter(lblTitle_.Name)
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

'// Some methods (GetColor and RaiseEvents)
Private Sub GetColor(ByRef isHover As Boolean)
    Dim clrFont As Long

    clrFont = IIf(isHover, oColor.clrActive, oColor.clrNormal)
    If lblTitle_.ForeColor <> clrFont Then
        lblTitle_.ForeColor = clrFont: txtInput_.ForeColor = clrFont
        lblULine_.BackColor = clrFont
    End If
End Sub
Private Sub OnChange(ByRef strValue As String)
    datValue = DateSerial(Val(Mid(strValue, 7, 4)), Val(Mid(strValue, 4, 2)), _
                          Val(Mid(strValue, 1, 2)))
    RaiseEvent Change
End Sub

'// Terminate
Private Sub Class_Terminate()
    If Not cTracker_ Is Nothing Then Set cTracker_ = Nothing
    If Not lblTitle_ Is Nothing Then Set lblTitle_ = Nothing
    If Not txtInput_ Is Nothing Then Set txtInput_ = Nothing
    If Not lblULine_ Is Nothing Then Set lblULine_ = Nothing
End Sub

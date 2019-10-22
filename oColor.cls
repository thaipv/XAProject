Option Explicit
Option Base 0

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
'Date:      18/06/2016
'Note:      https://vi.wikipedia.org/wiki/Danh_s%C3%A1ch_m%C3%A0u

Public Event ChangeColor()

Public Enum ThemeMode
    LightMode = 0
    DarkMode = 1
    'AutoMode = 2
End Enum

Public Enum ThemeColor
    XanhLam = 0
    XanhDuong = 1
    ThuyTinh = 2
    NonChuoi = 3
    NgocThach = 4
    LaCay = 5
    DoTuoi = 6
    DoTham = 7
    AnhDao = 8
    HongY = 9
    CamQuy = 10
    Tim = 11
    Man = 12
    Vang = 13
    HoPhach = 14
    DaCam = 15
End Enum

Private MainColor As ThemeColor
Private ModeColor As ThemeMode
Private arrColors As Variant

Private Sub Class_Initialize()
    ModeColor = LightMode   '// Get value from database
    MainColor = LaCay       '// Get value from database
    Call GetColors
End Sub

Friend Property Get clrFreeze() As Long:    clrFreeze = GetColor(0):    End Property
Friend Property Get clrNormal() As Long:    clrNormal = GetColor(1):    End Property
Friend Property Get clrBack() As Long:      clrBack = GetColor(2):      End Property
Friend Property Get clrDefault() As Long:   clrDefault = GetColor(3):   End Property
Friend Property Get clrInstant() As Long:   clrInstant = GetColor(4):   End Property
Friend Property Get clrActive() As Long:    clrActive = GetColor(5):    End Property
Friend Property Get clrWarning() As Long:   clrWarning = GetColor(6):   End Property
Friend Property Get Mode() As ThemeMode:    Mode = ModeColor:           End Property

Friend Sub ChangeTheme(ByRef bytTheme As ThemeColor)
    If MainColor <> bytTheme Then
        MainColor = bytTheme: Call GetColors
        RaiseEvent ChangeColor
    End If
End Sub
Friend Sub ChangeMode(ByRef bytMode As ThemeMode)
    If ModeColor <> bytMode Then
        ModeColor = bytMode: Call GetColors
        RaiseEvent ChangeColor
    End If
End Sub
Private Sub GetColors()
    If ModeColor = LightMode Then
        Select Case MainColor     'Freeze    Normal   vbWhite   Default   Instant   Active    Warning
        Case 0:  arrColors = Array(15790320, 3947580, 16777215, 16754297, 16740721, 16711680, 9471)  'Xanh lam
        Case 1:  arrColors = Array(15790320, 3947580, 16777215, 16762749, 16752157, 12611584, 9471)  'Xanh duong
        Case 2:  arrColors = Array(15790320, 3947580, 16777215, 16754297, 16734985, 10040064, 9471)  'Thuy tinh
        Case 3:  arrColors = Array(15790320, 3947580, 16777215, 9699273, 6422448, 65407, 9471)       'Nõn chuôi
        Case 4:  arrColors = Array(15790320, 3947580, 16777215, 14221203, 10944261, 7055360, 9471)   'Ngoc thach
        Case 5:  arrColors = Array(15790320, 3947580, 16777215, 12517257, 9436979, 5287936, 9471)    'Lá cây
        Case 6:  arrColors = Array(15790320, 3947580, 16777215, 9675519, 4941311, 9471, 65535)       'Do tuoi
        Case 7:  arrColors = Array(15790320, 3947580, 16777215, 10653685, 6373614, 3937500, 65535)   'Do tham
        Case 8:  arrColors = Array(15790320, 3947580, 16777215, 9996269, 8477157, 6500830, 65535)    'Anh dào
        Case 9:  arrColors = Array(15790320, 3947580, 16777215, 9996269, 6506979, 3808964, 65535)    'Hông y
        Case 10: arrColors = Array(15790320, 3947580, 16777215, 11437527, 9723594, 6697881, 65535)   'Câm quy
        Case 11: arrColors = Array(15790320, 3947580, 16777215, 16718313, 13762750, 10158219, 65535) 'Tím
        Case 12: arrColors = Array(15790320, 3947580, 16777215, 15073510, 10616994, 6684774, 65535)  'Mân
        Case 13: arrColors = Array(15790320, 3947580, 16777215, 7733247, 5767167, 65535, 9471)       'Vàng
        Case 14: arrColors = Array(15790320, 3947580, 16777215, 7199231, 4706303, 49151, 9471)       'Hô phách
        Case 15: arrColors = Array(15790320, 3947580, 16777215, 7199231, 2213375, 42495, 9471)       'Da cam
        Case Else: arrColors = Array(15790320, 3947580, 16777215, 12517257, 9436979, 5287936, 9471)  'Lá cây
        End Select
    Else 'ModeColor = DarkMode
        Select Case MainColor     'Freeze   Normal    vbBlack   Default  Instant   Active    Warning
        Case 0:  arrColors = Array(3947580, 11842740, 1513239, 16754297, 16740721, 16711680, 9471)   'Xanh lam
        Case 1:  arrColors = Array(3947580, 11842740, 1513239, 16762749, 16752157, 12611584, 9471)   'Xanh duong
        Case 2:  arrColors = Array(3947580, 11842740, 1513239, 16754297, 16734985, 10040064, 9471)   'Thuy tinh
        Case 3:  arrColors = Array(3947580, 11842740, 1513239, 9699273, 6422448, 65407, 9471)        'Nõn chuôi
        Case 4:  arrColors = Array(3947580, 11842740, 1513239, 14221203, 10944261, 7055360, 9471)    'Ngoc thach
        Case 5:  arrColors = Array(3947580, 11842740, 1513239, 12517257, 9436979, 5287936, 9471)     'Lá cây
        Case 6:  arrColors = Array(3947580, 11842740, 1513239, 9675519, 4941311, 9471, 65535)        'Do tuoi
        Case 7:  arrColors = Array(3947580, 11842740, 1513239, 10653685, 6373614, 3937500, 65535)    'Do tham
        Case 8:  arrColors = Array(3947580, 11842740, 1513239, 9996269, 8477157, 6500830, 65535)     'Anh dào
        Case 9:  arrColors = Array(3947580, 11842740, 1513239, 9996269, 6506979, 3808964, 65535)     'H?ng y
        Case 10: arrColors = Array(3947580, 11842740, 1513239, 11437527, 9723594, 6697881, 65535)    'Câm quy
        Case 11: arrColors = Array(3947580, 11842740, 1513239, 16718313, 13762750, 10158219, 65535)  'Tím
        Case 12: arrColors = Array(3947580, 11842740, 1513239, 15073510, 10616994, 6684774, 65535)   'Mân
        Case 13: arrColors = Array(3947580, 11842740, 1513239, 7733247, 5767167, 65535, 9471)        'Vàng
        Case 14: arrColors = Array(3947580, 11842740, 1513239, 7199231, 4706303, 49151, 9471)        'Hô phách
        Case 15: arrColors = Array(3947580, 11842740, 1513239, 7199231, 2213375, 42495, 9471)        'Da cam
        Case Else: arrColors = Array(3947580, 11842740, 1513239, 12517257, 9436979, 5287936, 9471)   'Lá cây
        End Select
    End If
End Sub
Private Function GetColor(ByRef bytIndex As Byte) As Long
    GetColor = arrColors(bytIndex)
End Function

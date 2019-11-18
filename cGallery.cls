VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGallery"
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
'Note:      Collection with events (Custom control)

Public Event Change(ByRef lngIterms As Long)
Private colSource As Collection

'// Initialize
Private Sub Class_Initialize()
    Set colSource = New Collection
End Sub

'// Some properties
Friend Property Get Count() As Long
    Count = colSource.Count
End Property
Friend Property Get Iterm(ByRef Index As Long) As Variant
    Iterm = colSource(Index)
End Property

'// Some methods
Friend Sub OnChange()
    RaiseEvent Change(Me.Count)
End Sub
Friend Sub Add(ByRef Iterm As Variant, Optional ByRef Key As Variant, _
               Optional ByRef Before As Variant, Optional ByRef After As Variant, _
               Optional ByRef isChange As Boolean = False)
    Call colSource.Add(Iterm, Key, Before, After)
    If isChange Then Call OnChange
End Sub
Friend Sub Remove(ByRef Index As Long, Optional ByRef isChange As Boolean = False)
    Call colSource.Remove(Index)
    If isChange Then Call OnChange
End Sub
Friend Sub Clear()
    Set colSource = Nothing
    Set colSource = New Collection
End Sub

'// Terminate
Private Sub Class_Terminate()
    Set colSource = Nothing
End Sub

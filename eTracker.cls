VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'Date:      #06/12/2017#
'Note:      Management center events of mouse

Public Event MouseTrack(strLeave As String, strHover As String)
Public ActiveControl As String
Friend Sub OnEnter(ByRef ControlName As String)
    If ActiveControl = ControlName Then Exit Sub
    RaiseEvent MouseTrack(ActiveControl, ControlName)
    ActiveControl = ControlName
End Sub

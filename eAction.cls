VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eAction"
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
'Date:      #06/12/2017#
'Note:      Central event monitoring activated controls

Public Event OnChange(strDeActive As String, strActive As String)
Private ActiveControl_ As String
Friend Property Get ActiveControl() As String
    ActiveControl = ActiveControl_
End Property
Friend Sub OnAction(ByRef ControlName As String)
    If ActiveControl_ = ControlName Then Exit Sub
    RaiseEvent OnChange(ActiveControl_, ControlName)
    ActiveControl_ = ControlName
End Sub

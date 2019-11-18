VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUnitCvt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Convert spacing unit"
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
'Date:      31/08/2019

#If VBA7 Then
    Private hDC As LongPtr
#Else
    Private hDC As Long
#End If
    
Private Const POINTSPERINCH = 72
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

'// Initialize
Private Sub Class_Initialize()
    hDC = GetDC(0)
End Sub

'// Conver points to pixels
Friend Function PointsToPixelsX(ByRef sngPoints As Single) As Long
    PointsToPixelsX = sngPoints * GetDeviceCaps(hDC, LOGPIXELSX) / POINTSPERINCH
End Function
Friend Function PointsToPixelsY(ByRef sngPoints As Single) As Long
    PointsToPixelsY = sngPoints * GetDeviceCaps(hDC, LOGPIXELSY) / POINTSPERINCH
End Function

'// Convert pixels to points
Friend Function PixelsToPointsX(ByRef lngPixels As Long) As Single
    PixelsToPointsX = lngPixels / GetDeviceCaps(hDC, LOGPIXELSX) * POINTSPERINCH
End Function
Friend Function PixelsToPointsY(ByRef lngPixels As Long) As Single
    PixelsToPointsY = lngPixels / GetDeviceCaps(hDC, LOGPIXELSY) * POINTSPERINCH
End Function

'// Convert centimeters to points   '// Convert points to centimeters
'// Convert centimeters to pixels   '// Convert pixels to centimeters

'// Terminate
Private Sub Class_Terminate()
    ReleaseDC 0, hDC
End Sub

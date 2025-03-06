VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objFileSys As Object

'// Initialize
Private Sub Class_Initialize()
    'Set objWsShell = CreateObject("WScript.Shell")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
End Sub

'// Get Desktop path
Friend Function GetDesktopPath() As String
    Dim objWsShell As Object
    
    Set objWsShell = CreateObject("WScript.Shell")
    GetDesktopPath = objWsShell.SpecialFolders("Desktop") & "\"
    
    Set objWsShell = Nothing
End Function

'// Get FileExists
Friend Function FileExists(ByRef strFilePath As String, ByRef strFileName As String) As Boolean
    FileExists = objFileSys.FileExists(strFilePath & strFileName)
End Function

'// Create a folder
Friend Function CreateFolder(ByRef strFolderPath As String) As Boolean
    On Error GoTo ErrHandle
        objFileSys.CreateFolder strFolderPath
    On Error GoTo 0
    
ErrHandle:
    CreateFolder = (Err.Number = 0)
End Function

'// Delete a folder
Friend Function DeleteFolder(ByRef strFolderPath As String) As Boolean
    On Error GoTo ErrHandle
        objFileSys.DeleteFolder strFolderPath
    On Error GoTo 0
    
ErrHandle:
    DeleteFolder = (Err.Number = 0)
End Function

'// Get parent folder name
Friend Function GetParentFolderName(ByRef strFolderPath As String) As String
    GetParentFolderName = objFileSys.GetParentFolderName(strFolderPath)
End Function

'// Bla bla bla bla ...

'// Teminate
Private Sub Class_Terminate()
    'Set objWsShell = Nothing
    Set objFileSys = Nothing
End Sub

Attribute VB_Name = "modHostJunk"
Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_Minimize = 6
Public Const SW_Maximize = 3
Public Const SW_Normal = 1

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const DefPort = 6394


'===Shell Anything===
'Desc: Acts like a dbl-Click on any file from the windows explorer
Public Sub Shellanything(Filetorun As String, Optional windowstate As Long = vbNormalFocus)
Call ShellExecute(0&, vbNullString, Filetorun, vbNullString, vbNullString, windowstate)
End Sub

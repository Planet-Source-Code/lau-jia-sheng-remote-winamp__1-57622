Attribute VB_Name = "Module1"
Public Declare Sub InitCommonControls Lib "comctl32" ()

Public Const DefPort = 6394
'================Start of Sys Icon=============='
    Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
    End Type
    
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const WM_MOUSEMOVE = &H200
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    
    Public Const WM_LBUTTONDBLCLK = &H203
    Public Const WM_LBUTTONDOWN = &H201
    Public Const WM_LBUTTONUP = &H202
    Public Const WM_RBUTTONDBLCLK = &H206
    Public Const WM_RBUTTONDOWN = &H204
    Public Const WM_RBUTTONUP = &H205
    Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Public taskicon As NOTIFYICONDATA, laststate As Integer
'===================End of Sys Icon============='


Sub load_icon(pichook As PictureBox, Icon As Picture, thetip As String)
On Error GoTo errhand
    taskicon.cbSize = Len(taskicon)
    taskicon.hwnd = pichook.hwnd
    taskicon.uId = 1&
    taskicon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    taskicon.ucallbackMessage = WM_MOUSEMOVE
    taskicon.hIcon = Icon
    taskicon.szTip = thetip & Chr$(0)
    Shell_NotifyIcon NIM_ADD, taskicon
Exit Sub
errhand:
MsgBox Err.Description & vbCrLf & "Has occured in Load_TrayIcon", , Err.Number
End Sub
Public Sub Unload_Icon(pichook As PictureBox)
On Error Resume Next
    taskicon.cbSize = Len(taskicon)
    taskicon.hwnd = pichook.hwnd
    taskicon.uId = 1&
    Shell_NotifyIcon NIM_DELETE, taskicon
End Sub

'===cms===
'Convert Minute to Seconds:  Makes seconds into easy to read hh:mm:ss
Public Function cms(seconds As Single) As String
    'First Convert into Minutes, and find left over seconds
    On Error GoTo errhand
    Dim Min As Long, Secs As Byte
    If seconds <= 0 Then GoTo errhand
    Min = seconds \ 60
    Secs = seconds Mod 60 'the left over seconds
    cms = CStr(Min) & ":" & CStr(Format(CStr(Secs), "00"))
    Exit Function
errhand:
    cms = "0:00"
End Function

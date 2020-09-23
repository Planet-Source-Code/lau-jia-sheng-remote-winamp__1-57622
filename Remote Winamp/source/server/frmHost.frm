VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmHost 
   BorderStyle     =   0  'None
   Caption         =   "Remote Winamp Server"
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   -195
   ClientWidth     =   1560
   Icon            =   "frmHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock WinsockCommands 
      Index           =   0
      Left            =   540
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Text            =   "C:\"
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2220
      TabIndex        =   2
      Top             =   1500
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   2340
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock WinsockListen 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WinsockCommands 
      Index           =   1
      Left            =   540
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WinsockCommands 
      Index           =   2
      Left            =   540
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileItems() As String, fileitemindex As Integer
Private Function RemoveParent(ByVal File As String) As String
On Error GoTo errhand
Dim t As Long
If Right(File, 1) = "\" Then File = Mid(File, 1, Len(File) - 1)
For t = Len(File) To 1 Step -1
If Mid(File, t, 1) = "\" Then
    RemoveParent = Mid(File, t + 1)
    Exit Function
End If
Next t
errhand:
End Function
Private Function RemoveFileName(ByVal File As String) As String
On Error GoTo errhand
Dim t As Long
If Right(File, 1) = "\" Then File = Mid(File, 1, Len(File) - 1)
For t = Len(File) To 1 Step -1
If Mid(File, t, 1) = "\" Then
    RemoveFileName = Left(File, t)
    Exit Function
End If
Next t
errhand:
End Function


Function Search(Path As String, spec As String, Optional asArray As Boolean = True) As String()
On Error GoTo errhand
Dim x() As String, t As Integer, i As Integer
Dir1.Path = Path
File1.Path = Path
File1.Pattern = spec
File1.Archive = True: File1.ReadOnly = True: File1.System = True: File1.Hidden = True
DoEvents
If Dir1.ListCount = 0 And File1.ListCount = 0 Then
If Not asArray Then
ReDim x(0)
x(0) = Replace(Dir1.Path & "\", "\\", "\") & "|" & "..\"
Else
ReDim x(1)
x(0) = Replace(Dir1.Path & "\", "\\", "\")
x(1) = "..\"
End If
Search = x
Exit Function
End If
If Not asArray Then
ReDim x(0)
x(0) = Replace(Dir1.Path & "\", "\\", "\") & "|" & "..\"
Else
ReDim x(Dir1.ListCount + File1.ListCount + 2)
x(0) = Replace(Dir1.Path & "\", "\\", "\")
x(1) = "..\"
End If
For t = 0 To Dir1.ListCount - 1
    If Not asArray Then
    x(0) = x(0) & "|" & RemoveParent(Dir1.List(t)) & "\"
    Else
    x(t + 2) = RemoveParent(Dir1.List(t)) & "\"
    End If
Next t
t = t + 2
For i = 0 To File1.ListCount - 1
    If Not asArray Then
    x(0) = x(0) & "|" & File1.List(i)
    Else
    x(t) = File1.List(i)
    End If
    t = t + 1
Next i
Search = x
Exit Function
errhand:
If Not asArray Then
    ReDim x(0)
    x(0) = Err.Number & "|" & Err.Description & "|" & Path & "|" & spec
Else
    ReDim x(3)
    x(0) = Err.Number
    x(1) = Err.Description
    x(2) = Path
    x(3) = spec
End If
Search = x
End Function

Private Sub Command1_Click()
On Error GoTo errhand
Dim x() As String, i As Integer
x = Search(Text1, "*.*")
List1.Clear
Label1 = x(0)
For i = 1 To UBound(x)
List1.AddItem x(i)
Next i
errhand:
End Sub

Private Sub Form_Load()
On Error GoTo errhand
Call WriteRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Service", App.Path & "\" & App.EXEName & ".exe")
AutoLoadWinAmp = True
winamppath = GetSetting("RWServer", "REMOTEWINAMPHOST", "Winamppath", "")
FindWinampPath
errhand:
WinsockListen.LocalPort = DefPort
WinsockListen.Listen
App.TaskVisible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'//Uncomment to prevent the user from disbaleding the
'//turning off this program

'If UnloadMode = vbAppTaskManager Then
'WinsockCommands(0).Close: WinsockCommands(1).Close: WinsockCommands(2).Close
'WinsockListen.Close
'Shell Replace(App.Path & "\" & App.EXEName, "\\", "\")
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "RWServer", "REMOTEWINAMPHOST", "Winamppath", winamppath
End
End Sub

Private Sub List1_Click()
On Error GoTo errhand
If Right(List1.Text, 1) = "\" Then Text1 = Label1 & List1.Text
errhand:
End Sub

Private Sub WinsockCommands_Close(Index As Integer)
WinsockCommands(Index).Close
WinsockCommands(Index).RemotePort = 0: WinsockCommands(Index).LocalPort = 0
End Sub

Private Sub WinsockListen_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer
For i = 0 To 2
If WinsockCommands(i).State = sckClosed Then
    WinsockCommands(i).Accept requestID
    Debug.Print "connected to:" & WinsockCommands(i).RemoteHostIP
    Exit Sub
End If
Next i
End Sub
Private Sub WinsockCommands_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo errhand
Dim newdata As String, Linestart As Long, linedata As String
Static lastData(2) As String
WinsockCommands(Index).PeekData newdata, vbString
If InStr(1, newdata, Chr(1)) = 0 Then
    Exit Sub
End If
WinsockCommands(Index).GetData newdata, vbString
'//Logdata newdata
If lastData(Index) <> "" Then newdata = lastData(Index) & newdata
again: Linestart = InStr(1, newdata, Chr(1))
If Linestart <> 0 Then
linedata = Mid(newdata, 1, Linestart - 1)
newdata = Mid(newdata, Linestart + 1)
ProcessData linedata, Index
Else
lastData(Index) = newdata
newdata = ""
End If
If newdata <> "" And newdata <> Chr(1) Then GoTo again
Exit Sub
errhand:
MsgBox Err.Number & vbCrLf & Err.Description, , "Winsock3_DataArrival"
End Sub
Sub SendData(data As String, Index As Integer)
On Error GoTo errhand
If WinsockCommands(Index).State = sckConnected Then
WinsockCommands(Index).SendData data & Chr(1)
End If
errhand:
End Sub

Private Sub ProcessData(linedata As String, Index As Integer)
Dim datatype As String, data As String
Dim t As Integer, path_pls As String
Dim trake As Long, times As Long, PLSLNG As Long, trale As String
datatype = Left(linedata, 4)
data = Mid(linedata, 5)
AutoLoadWinAmp = True
Select Case UCase(datatype)
Case "PLAY"
WM_DO (PLAY_TRACK)
Case "STOP"
WM_DO (STOP_TRACK)
Case "NEXT"
WM_DO (NEXT_TRACK)
Case "PREV"
WM_DO (PREV_TRACK)
Case "HALT"
WM_DO (PAUSE_TRACK)
Case "VOLU"
WM_SetVolume CByte(data)
Case "BALN"
WM_SetPanning CByte(data)
Case "SHUF"
WM_DO TOGGLE_SHUFFLE
Case "REPE"
WM_DO TOGGLE_REPEAT
Case "AONT"
WM_DO TOGGLE_ALWAYS_ON_TOP
Case "DBLM"
WM_DO TOGGLE_DOUBLESIZE
Case "CLER"
WM_GET CLEAR_PLAYLIST
Case "VISA"
WM_DO START_VIS
Case "GBEG"
WM_DO GO_BEGINNING_PLAYLIST
Case "GEND"
WM_DO GO_END_PLAYLIST
Case "BRWD" 'Browse Directory
Dim Path As String, spec As String
Dim items() As String
'data: Directory info to send.
'e.g. C:\My music\*.*
'e.g. C:\My Music\*.mp3
Path = RemoveFileName(data)
spec = RemoveParent(data)
items = Search(Path, spec, False) 'no as array to make it easier to send
SendData "DIRI" & items(0), Index
Case "REQU" 'Browse Directory
Load frmMain
frmMain.txtHost = WinsockCommands(Index).RemoteHostIP
frmMain.txtPath = data
frmMain.tmrSend.Enabled = True
Case "LFLE"
'this host machine local path!
'data: "D:\CD\Pls\JamesList.pls"
If Dir(winamppath) <> "" Then
If Right(data, 1) = "\" Then data = Mid(data, 1, Len(data) - 1)
Shell winamppath & " " & """" & data & """"
SendData "CLFL1", Index
Else
SendData "CLFL0", Index
End If
Case "AFLE"
'this host machine local path!
'data: "D:\CD\Pls\JamesList.pls"
If Dir(winamppath) <> "" Then
If Right(data, 1) = "\" Then data = Mid(data, 1, Len(data) - 1)
Shell winamppath & " /ADD " & """" & data & """"
SendData "CAFL1", Index
Else
SendData "CAFL0", Index
End If
Case "WAMP"
'the path on this host machine local HD
winamppath = data
Case "CLSW"
WM_DO WINAMP_EXIT
Case "CLOS"
Unload Me
Case "SNIN"
AutoLoadWinAmp = False
times = WM_GET(SONG_POSITION)
trake = WM_GET(PLAYLIST_POSITION)
PLSLNG = WM_GET(PLAYLIST_LENGTH)
trale = WM_GET(SONG_TITLE)
SendData "SNIN" & CStr(times) & Chr(3) & CStr(trake) & Chr(3) & CStr(PLSLNG) & Chr(3) & CStr(trale), Index
AutoLoadWinAmp = True
Case "MINW"
ShowWindow WinampID, SW_Minimize
Case "RESW"
ShowWindow WinampID, SW_Normal
Case "FIND"
'data=D:\cd\|*.mp3|0=f1=t
FileItems = Split(data, "|", 3)
FileItems = Searchfiles(FileItems(0), FileItems(1), CBool(FileItems(2)))
For t = 0 To UBound(FileItems)
    If FileItems(t) <> "" Then
    trale = trale & FileItems(t) & vbCrLf
    End If
Next t
SendData "FIND" & trale, Index
Case "GPLS"
WM_GET WRITE_CURR_PLAYLIST
path_pls = Replace(RemoveFileName(FindWinampPath) & "\winamp.m3u", "\\", "\")
trale = MakePLSString(path_pls)
SendData "GPLS" & trale, Index
Case "GOTO"
WM_SetPlaylistPos CInt(data)
WM_DO (PLAY_TRACK)
End Select
End Sub
Function MakePLSString(Path As String) As String
Dim t As Integer
Dim FF As Integer
Dim linedata As String
FF = FreeFile
If Dir(Path) <> "" Then
Open Path For Input As #FF
    Do While Not EOF(FF)
        Line Input #FF, linedata
        If Left(linedata, 1) <> "#" Then
            MakePLSString = MakePLSString & linedata & vbCrLf
        End If
    Loop
Close #FF
Else
MakePLSString = ""
End If
End Function

Private Function Searchfiles(Path As String, Filespec As String, recursive As Boolean) As String()
  Dim asfiles As Variant
  Dim lLoop As Long
  Dim lCount As Long
  Dim bResult As Boolean
  Dim Clsfind As New clsGetFiles

  Clsfind.Path = Path 'UNC Paths are supported
  Clsfind.Filespec = Filespec  'Wild Cards are also supported
  Clsfind.recursive = recursive
  
  bResult = Clsfind.FindAll(asfiles)

  If VarType(asfiles) = (vbArray + vbString) Then
    Searchfiles = asfiles
  Else
    ReDim asfiles(0)
    asfiles(0) = "No Files Found"
    Searchfiles = asfiles
  End If
End Function

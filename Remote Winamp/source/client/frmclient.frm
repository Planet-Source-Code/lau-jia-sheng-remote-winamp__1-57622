VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmclient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status: Disconnected"
   ClientHeight    =   7350
   ClientLeft      =   300
   ClientTop       =   2025
   ClientWidth     =   8295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmclient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPauseTrans 
      Caption         =   "Pause Transmitting Signals (To prevent detection)"
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   1170
      Width           =   4635
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   315
      Left            =   6540
      TabIndex        =   29
      Top             =   6600
      Width           =   1395
   End
   Begin VB.CommandButton cmdIPScan 
      Caption         =   "Scan for LAN Users in Local Network"
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   800
      Width           =   4635
   End
   Begin ComctlLib.Slider sldVol 
      Height          =   255
      Left            =   1620
      TabIndex        =   6
      Top             =   2280
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   450
      _Version        =   327682
      LargeChange     =   10
      Max             =   255
      SelStart        =   127
      TickStyle       =   3
      Value           =   127
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   3510
      TabIndex        =   25
      Top             =   6600
      Width           =   915
   End
   Begin VB.CommandButton cmdshowPlaylist 
      Caption         =   "PL"
      Height          =   375
      Left            =   7570
      TabIndex        =   33
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Search"
      Height          =   375
      Left            =   6540
      TabIndex        =   32
      Top             =   5760
      Width           =   1000
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   315
      Left            =   4460
      TabIndex        =   27
      Top             =   6600
      Width           =   915
   End
   Begin VB.CommandButton cmdaddfile 
      Caption         =   "Add to PL"
      Height          =   315
      Left            =   5400
      TabIndex        =   28
      Top             =   6600
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Modify Remote Winamp Path"
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Connect/Disconnect"
      Top             =   420
      Width           =   1755
   End
   Begin VB.CommandButton cmdTerminateCL 
      Caption         =   "Terminate RW Client"
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdCloseRW 
      Caption         =   "Terminate RW Server"
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   195
      Width           =   2535
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Previous Track"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Next Track"
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   14
      Top             =   2955
      Width           =   1335
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Play"
      Height          =   375
      Index           =   2
      Left            =   1710
      TabIndex        =   11
      Top             =   3120
      Width           =   650
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Pause"
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   12
      Top             =   3120
      Width           =   650
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Stop"
      Height          =   375
      Index           =   4
      Left            =   3090
      TabIndex        =   13
      Top             =   3120
      Width           =   650
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit Remote Winamp"
      Height          =   375
      Left            =   5520
      TabIndex        =   23
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   " Built-in Browser (Search for files remotely) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   120
      TabIndex        =   36
      Top             =   3960
      Width           =   8055
      Begin VB.ComboBox cmbdblclick 
         Height          =   315
         ItemData        =   "frmclient.frx":08CA
         Left            =   6400
         List            =   "frmclient.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbFiletype 
         Height          =   315
         ItemData        =   "frmclient.frx":08F1
         Left            =   6400
         List            =   "frmclient.frx":090D
         TabIndex        =   31
         Text            =   "*.*"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtselfile 
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         Top             =   2660
         Width           =   1995
      End
      Begin VB.ListBox lstfiles 
         Height          =   2205
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Browse for:"
         Height          =   255
         Left            =   6360
         TabIndex        =   42
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Double-click Action:"
         Height          =   255
         Left            =   6360
         TabIndex        =   41
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2700
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   1080
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "First Track"
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   2955
      Width           =   1335
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Last Track"
      Height          =   375
      Index           =   10
      Left            =   3840
      TabIndex        =   15
      ToolTipText     =   "Last Track"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Clear Remote Winamp Playlist"
      Height          =   375
      Index           =   7
      Left            =   5520
      TabIndex        =   19
      Top             =   1605
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Interval        =   133
      Left            =   5040
      Top             =   1560
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Minimise Remote Winamp"
      Height          =   375
      Left            =   5520
      TabIndex        =   21
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Restore Remote Winamp"
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   3045
      Width           =   2535
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Remote Winamp Visualisation"
      Height          =   375
      Index           =   8
      Left            =   5520
      TabIndex        =   20
      Top             =   2010
      Width           =   2535
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Shuffle"
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdActions 
      Caption         =   "Repeat"
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2325
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Center the Balance"
      Height          =   785
      Left            =   3960
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   " Configure Remote Connection "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Text            =   "127.0.0.1"
         Top             =   320
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.PictureBox pictray 
      Height          =   315
      Left            =   4560
      ScaleHeight     =   255
      ScaleWidth      =   435
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5040
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.Slider sldBal 
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   2280
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      _Version        =   327682
      LargeChange     =   10
      Max             =   255
      SelStart        =   127
      TickStyle       =   3
      Value           =   127
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00946934&
      Height          =   255
      Left            =   3000
      TabIndex        =   38
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Left            =   1650
      TabIndex        =   37
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim winamppath As String
Dim PlNum As Long, PlTot As Long, songtime As Long, SongName As String
Private Function RemoveParent(ByVal File As String) As String
Dim t As Long
If Right(File, 1) = "\" Then File = Mid(File, 1, Len(File) - 1)
For t = Len(File) To 1 Step -1
If Mid(File, t, 1) = "\" Then
    RemoveParent = Mid(File, t + 1)
    Exit Function
End If
Next t
End Function
Private Function RemoveFileName(ByVal File As String) As String
Dim t As Long
If Right(File, 1) = "\" Then File = Mid(File, 1, Len(File) - 1)
For t = Len(File) To 1 Step -1
If Mid(File, t, 1) = "\" Then
    RemoveFileName = Left(File, t)
    Exit Function
End If
Next t
End Function


Private Sub cmdActions_Click(Index As Integer)
Select Case Index
Case 0
SendData "PREV"
Case 1
SendData "NEXT"
Case 2
SendData "PLAY"
Case 3
SendData "HALT"
Case 4
SendData "STOP"
Case 5
SendData "REPE"
Case 6
SendData "SHUF"
Case 7
SendData "CLER"
Case 8
SendData "VISA"
Case 9
SendData "GBEG"
Case 10
SendData "GEND"

End Select
End Sub

Private Sub cmdaddfile_Click()
SendData "AFLE" & txtselfile
End Sub

Private Sub cmdCloseRW_Click()
SendData "CLOS"
Timer2.Enabled = False
Me.Caption = "Status: Disconnected"
'End
End Sub

Private Sub cmdDownload_Click()
frmMain.Show
frmMain.txtfilename = lstfiles.Text
SendData "REQU" & txtselfile
End Sub

Private Sub cmdfind_Click()
frmfindfiles.FillinFields txtselfile
'frmfindfiles.Visible = True
frmfindfiles.Show , Me
End Sub

Private Sub cmdIPScan_Click()
frmIpScanner.Show , Me
End Sub

Private Sub cmdPauseTrans_Click()
If Timer2.Enabled = True Then
Timer2.Enabled = False
cmdPauseTrans.Caption = "Resume Tansmitting Signals without Disconnection"
Else
Timer2.Enabled = True
cmdPauseTrans.Caption = "Pause Tansmitting Signals without Disconnection"
End If
End Sub

Private Sub cmdshowPlaylist_Click()
SendData "GPLS"
'frmplaylist.Visible = True
frmplaylist.Show , Me
End Sub

Private Sub cmdTerminateCL_Click()
Unload Me
End Sub

Private Sub Command10_Click()
SendData "RESW"
End Sub

Private Sub cmdReplace_Click()
SendData "LFLE" & txtselfile
End Sub

Private Sub Command3_Click()
Dim X As String
X = InputBox("Modify winamp's path in the server-side?" & vbCrLf & "Press the cancel button if winamp is already working!" & vbCrLf & "Setting the wrong path may cause the server-computer to crash!" & vbCrLf & "" & vbCrLf & "(Use the built-in browser to find the winamp.exe file)", "Modify Server-Side Winamp Path", winamppath)
If X = "" Then Exit Sub
winamppath = X
SendData "WAMP" & winamppath
End Sub

Private Sub Command4_Click()
SendData "CLSW"
End Sub

Private Sub cmdBrowse_Click()
If Right(txtselfile, 1) <> "\" Then
    MsgBox "Add a \ behind the path you entered!", vbCritical, "Element missing!"
    Exit Sub
End If
SendData "BRWD" & Replace(txtselfile & "\" & cmbFiletype.Text, "\\", "\")
End Sub

Private Sub Command6_Click()
If Command6.Caption = "&Connect" Then
Timer2.Enabled = False
Winsock1.Close
Timer2.Enabled = True
Winsock1.Connect txtIP, DefPort
Else
Timer2.Enabled = False
Winsock1.Close
Winsock1.RemoteHost = ""
Winsock1.RemotePort = 0
Winsock1.LocalPort = 0
Me.Caption = "Status: Disconnected"
Command6.Caption = "&Connect"
PlNum = "0"
SongName = "??"
songtime = "0"
Beep
End If
End Sub

Private Sub Command8_Click()
sldBal.Value = 128
End Sub

Private Sub Command9_Click()
SendData "MINW"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbShiftMask Or vbAltMask Or vbCtrlMask Then
    If KeyCode = vbKeyC Then
        SendData "CHSS"
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
winamppath = GetSetting("RWClient", "REMOTEWINAMP", "WinampPath", "")
txtIP = GetSetting("RWClient", "REMOTEWINAMP", "IP", "127.0.0.1")
cmbdblclick.ListIndex = GetSetting("RWClient", "REMOTEWINAMP", "DoubleClick", 0)
load_icon pictray, Me.Icon, "Remote Winamp - Client"
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "RWClient", "REMOTEWINAMP", "WinampPath", winamppath
SaveSetting "RWClient", "REMOTEWINAMP", "IP", txtIP
SaveSetting "RWClient", "REMOTEWINAMP", "DoubleClick", cmbdblclick.ListIndex
Unload_Icon pictray
End
End Sub

Private Sub lstfiles_Click()
If Mid(lstfiles.Tag & lstfiles.List(lstfiles.ListIndex), 2) = ":\..\" Then
txtselfile = ""
Else
txtselfile = lstfiles.Tag & lstfiles.List(lstfiles.ListIndex)
End If
End Sub

Private Sub lstfiles_DblClick()
If Mid(lstfiles.Tag & lstfiles.List(lstfiles.ListIndex), 2) = ":\..\" Then
txtselfile = ""
Else
txtselfile = lstfiles.Tag & lstfiles.List(lstfiles.ListIndex)
End If
Select Case cmbdblclick.ListIndex
Case 0  'Browse
If Right(txtselfile, 1) <> "\" Then
    MsgBox "Invalid Directory!", vbCritical, "Error"
    Exit Sub
End If
SendData "BRWD" & Replace(txtselfile & "\" & cmbFiletype.Text, "\\", "\")
Case 1 'Add
cmdaddfile_Click
Case 2 'Replace
cmdReplace_Click
End Select
End Sub

Private Sub pictray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhand
Dim msg As Long
    msg = X / Screen.TwipsPerPixelX
        Select Case msg
            Case WM_LBUTTONDBLCLK:
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP
            Me.Visible = Not Me.Visible
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
            Me.Visible = Not Me.Visible
        End Select
errhand:
End Sub

Private Sub sldVol_Change()
Call sldVol_Scroll
End Sub

Private Sub sldVol_Scroll()
SendData "VOLU" & sldVol.Value
End Sub
Private Sub sldBal_Change()
Call sldBal_Scroll
End Sub

Private Sub sldBal_Scroll()
'Dim prcnt As Long
SendData "BALN" & sldBal.Value
'prcnt = Int((sldBal.Value - 127) / 1.27)
'If prcnt = 0 Then
'    Timer1.Enabled = True
'Else
'    lrstatus = ""
'    Timer1.Enabled = False
'End If
End Sub


'Private Sub Timer1_Timer()
'If lrstatus.Caption <> "" Then
'lrstatus.Caption = ""
'Timer1.Enabled = False
'End If
'End Sub

Private Sub Timer2_Timer()
Static times As Byte
times = times + 1
If times = 5 Then
SendData ("SNIN")
times = 0
Else
If songtime > 0 Then
songtime = songtime
'Me.Caption = (PlNum + 1) & "\" & PlTot & " " & cms(CSng(songtime / 1000)) & "   " & SongName
Me.Caption = (PlNum + 1) & ". " & SongName & " - [" & cms(CSng(songtime / 1000)) & "]"
End If
End If
End Sub

Private Sub txtIP_GotFocus()
Command6.Default = True
End Sub

Private Sub txtselfile_GotFocus()
cmdBrowse.Default = True
End Sub

Private Sub Winsock1_Close()
Timer2.Enabled = False
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
Timer2.Enabled = True
Me.Caption = "Status: Connected"
Command6.Caption = "Dis&connect"
Beep
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errhand
Dim newdata As String, Linestart As Long, linedata As String
Static lastData As String
Winsock1.PeekData newdata, vbString
If InStr(1, newdata, Chr(1)) = 0 Then
    Exit Sub
End If
Winsock1.GetData newdata, vbString
'//Logdata newdata
again: Linestart = InStr(1, newdata, Chr(1))
If Linestart <> 0 Then
linedata = Mid(newdata, 1, Linestart - 1)
newdata = Mid(newdata, Linestart + 1)
ProcessData linedata
Else
lastData = newdata
newdata = ""
End If
If newdata <> "" And newdata <> Chr(1) Then GoTo again
Exit Sub
errhand:
MsgBox Err.Number & vbCrLf & Err.Description, , "Winsock1_DataArrival"
End Sub
Private Sub ProcessData(linedata As String)
Dim datatype As String, data As String
Dim t As Integer, shortname As String
Dim pathItems() As String
datatype = Left(linedata, 4)
data = Mid(linedata, 5)
Select Case UCase(datatype)
Case "DIRI" 'directory info
pathItems = Split(data, "|")
lstfiles.Clear
If IsNumeric(pathItems(0)) = True Then
lstfiles.Tag = ""
lstfiles.AddItem "ERROR#" & pathItems(0)
lstfiles.AddItem pathItems(1)
lstfiles.AddItem pathItems(2)
lstfiles.AddItem pathItems(3)
Else
lstfiles.Tag = pathItems(0)
For t = 1 To UBound(pathItems)
    lstfiles.AddItem pathItems(t)
Next t
End If
Case "SNIN"
'data=SongTime|TrkNum|#Trk|SongTitle
'where | is CHr(3)
pathItems = Split(data, Chr(3), 4)
'Me.Caption = (pathItems(1) + 1) & "\" & pathItems(2) & " " & cms(CSng(pathItems(0) / 1000)) & "   " & pathItems(3)
PlNum = pathItems(1)
PlTot = pathItems(2)
songtime = pathItems(0)
SongName = pathItems(3)
Case "CLRP"

Case "PLSE"

Case "FIND"
'data: filename(LB)Filename(LB)...
pathItems = Split(data, vbCrLf)
For t = 0 To UBound(pathItems)
frmfindfiles.lstfoundfiles.AddItem pathItems(t), 0
Next t
Case "GPLS"
'Data=Playlist Entry#1(LB)PLS Entry#2(LB)...
pathItems = Split(data, vbCrLf)
frmplaylist.lstpls.Clear
For t = 0 To UBound(pathItems)
    If pathItems(t) <> "" Then
        shortname = RemoveParent(pathItems(t))
        If t + 1 < 10 Then
            pathItems(t) = "   " & t + 1 & ".  " & pathItems(t)
        ElseIf t + 1 < 100 Then
            pathItems(t) = "  " & t + 1 & ".  " & pathItems(t)
        ElseIf t + 1 < 1000 Then
            pathItems(t) = " " & t + 1 & ".  " & pathItems(t)
        Else
            pathItems(t) = t + 1 & ".  " & pathItems(t)
        End If
        frmplaylist.lstpls.AddItem pathItems(t)
    
        If t + 1 < 10 Then
            pathItems(t) = "   " & t + 1 & ".  " & shortname
        ElseIf t + 1 < 100 Then
            pathItems(t) = "  " & t + 1 & ".  " & shortname
        ElseIf t + 1 < 1000 Then
            pathItems(t) = " " & t + 1 & ".  " & shortname
        Else
            pathItems(t) = t + 1 & ".  " & shortname
        End If
        frmplaylist.lstshortpls.AddItem pathItems(t)
    End If
Next t
End Select

End Sub


Sub SendData(data As String)
If Winsock1.State = sckConnected Then
Winsock1.SendData data & Chr(1)
End If
End Sub


Private Sub Form_Initialize()
    InitCommonControls
End Sub


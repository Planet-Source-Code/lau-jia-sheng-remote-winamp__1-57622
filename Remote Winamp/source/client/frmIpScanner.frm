VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIpScanner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan for LAN Users in Local Network"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIpScanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   2880
      Top             =   3240
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   3840
      Top             =   3240
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   2880
      Top             =   2280
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   3840
      Top             =   2280
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   100
      Left            =   3840
      Top             =   1800
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   100
      Left            =   2880
      Top             =   3720
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   100
      Left            =   3840
      Top             =   3720
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   100
      Left            =   2880
      Top             =   2760
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   100
      Left            =   3840
      Top             =   2760
   End
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   100
      Left            =   2880
      Top             =   1800
   End
   Begin VB.CommandButton cmdStopScan 
      Caption         =   "Stop Scan"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   2535
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdStartScan 
      Caption         =   "Start Scan"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   3360
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   1
      Left            =   3360
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   2
      Left            =   2880
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   3
      Left            =   3360
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   4
      Left            =   3840
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   5
      Left            =   3360
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   6
      Left            =   3360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   7
      Left            =   3360
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   8
      Left            =   3120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   9
      Left            =   3600
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstIpes 
      Height          =   5325
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Double-click on an item to connect"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame frasettings 
      Caption         =   " Settings "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2535
      Begin VB.TextBox txtport 
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Text            =   "139"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtFirst3 
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Text            =   "192.168.0."
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Text            =   "100"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Program port is 6394)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Default Lan port is 139)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Port to scan:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Speed of scan, 1 per x ms)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(xxx.xxx.xxx.1-255)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address:"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblInvetval 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Timer Interval:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmIpScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Timer As Integer
Private IP As Integer

Private Sub cmdStartScan_Click()
Form_Load
tmrScan(Timer).Enabled = True
cmdStartScan.Enabled = False
cmdStopScan.Enabled = True
lstIpes.Clear
End Sub

Private Sub cmdStopScan_Click()
Dim I As Integer
I = 0
Do While I <= 9
  tmrScan(I).Enabled = False
  I = I + 1
Loop
Form_Load
End Sub

Private Sub Form_Load()
IP = 0
Timer = 0
cmdStartScan.Enabled = True
cmdStopScan.Enabled = False
End Sub

Private Sub lstIpes_DblClick()
frmclient.txtIP = lstIpes.List(lstIpes.ListIndex)
Unload Me
End Sub

Private Sub lstIpes_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyExecute Or KeyCode = vbKeyReturn Then
frmclient.txtIP = lstIpes.List(lstIpes.ListIndex)
Unload Me
End If
End Sub

Private Sub tmrScan_Timer(Index As Integer)
On Error GoTo Err:
Dim IPTS As String
If IP = 255 Then
  MsgBox "Done Scanning, found a total of " & lstIpes.ListCount & " IP(s) on the LAN network.", vbInformation, "Task Completed"
  txtIP = "Done Scanning: " & lstIpes.ListCount & " match(es)"
  IP = 0
  Do While IP <= 9
    tmrScan(IP).Enabled = False
    IP = IP + 1
  Loop
  Form_Load
  Exit Sub
End If
IP = IP + 1
ws(Timer).Close
IPTS = txtFirst3 & IP
txtIP = "Now Scanning: " & IPTS
ws(Timer).Connect IPTS, txtport
Timer = Timer + 1
If Timer > 9 Then Timer = 0
tmrScan(Timer).Enabled = True
tmrScan(Index).Enabled = False
Err:
Exit Sub
End Sub

Private Sub txtinterval_Change()
Dim I As Integer
If txtInterval = "" Then txtInterval = 1: txtInterval.SetFocus
If IsNumeric(txtInverval) = True Then
  I = 0
  Do While I <= 9
    tmrScan(I).Interval = txtInterval
    I = I + 1
  Loop
Else
MsgBox "Error: Interval could only be a number in ms, resetting.", vbCritical, "Error"
txtInterval = 100
End If
End Sub

Private Sub ws_Connect(Index As Integer)
lstIpes.AddItem ws(Index).RemoteHostIP
ws(Index).Close
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
ws(Index).Close
End Sub

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download a remote file (Silent)"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " Downloading Properties "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.TextBox txtfilename 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   6855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saving file as:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label lblDirFolder 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   6855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Download folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSWinsockLib.Winsock RecInfo 
      Left            =   3360
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9010
   End
   Begin MSWinsockLib.Winsock RecData 
      Left            =   3360
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   4050
   End
   Begin MSWinsockLib.Winsock SendInfo 
      Left            =   3960
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9010
   End
   Begin MSWinsockLib.Winsock SendData 
      Left            =   3960
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4050
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4800
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status: Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   8535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type FileData1
FileName As String
FileSize As Long
End Type

Private Type RecFile1
FileName As String
FileSize As Long
End Type

Dim FileData As FileData1

Dim RecFile As RecFile1

Dim bCon As Boolean

Dim CurByte As Long
Dim SendTotal As Long

Dim FileNum As Integer

Private Const sHeader As String = "Z{L"
Private Const sDelim As String = "a<  "

Private Const PacketSize As Long = 4096

Private Sub Form_Load()
FileNum = FreeFile
On Error Resume Next
RecInfo.Listen
End Sub

Private Sub RecData_Close()
Close #FileNum
lblStatus = "Status : File Received."
RecData.Close
RecInfo.Close
On Error Resume Next
RecInfo.Listen
End Sub

Private Sub RecData_Connect()
Call MakeRecDir
Open App.Path & "\Received Files\" & txtfilename For Binary Access Write As #FileNum
lblDirFolder = App.Path & "\Received Files\"
lblStatus = "Status : Receiving File . . ."
End Sub

Private Sub MakeRecDir()
On Error Resume Next
MkDir App.Path & "\Received Files"
End Sub

Private Sub RecData_DataArrival(ByVal bytesTotal As Long)
Dim FileDat As String: FileDat = Empty
RecData.GetData FileDat
Put #FileNum, , FileDat
End Sub

Private Sub RecData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lblStatus = "Status : Unable to Receive File."
End Sub

Private Sub RecInfo_Close()
On Error Resume Next
RecInfo.Close
RecInfo.Listen
End Sub

Private Sub RecInfo_ConnectionRequest(ByVal requestID As Long)
RecInfo.Close
RecInfo.Accept requestID
End Sub

Private Sub RecInfo_DataArrival(ByVal bytesTotal As Long)
Dim sRecData As String: sRecData = Empty
Dim sBuff() As String
RecInfo.GetData sRecData
Select Case Mid(sRecData, 12, 1)
Case "F"
    sBuff() = Split(sRecData, sDelim)
    Dim sRetPack As String, iRep As Integer
    iRep = MsgBox("You requested " & "'" & txtfilename & "' from " & RecInfo.RemoteHostIP & vbCrLf & "Would you like to accept the " & sBuff(3) & " bytes file?", vbQuestion + vbYesNo, "Download remote files")
        If iRep = vbNo Then
            sRetPack = sHeader & sDelim & "F" & sDelim & "Denied"
            RecInfo.SendData sRetPack
        Else
            
            RecFile.FileName = sBuff(2)
            RecFile.FileSize = CLng(Val(sBuff(3)))
            sRetPack = sHeader & sDelim & "F" & sDelim & "Accepted"
            lblStatus = "Status : Negotiating Transfer . . ."
            RecInfo.SendData sRetPack
        End If
    sRetPack = Empty: iRep = 0

Case "R"
    RecData.Close
    RecData.Connect RecInfo.RemoteHostIP, 4050
    
End Select
End Sub

Private Sub SendData_Close()
On Error Resume Next
SendData.Close
SendData.Listen
End Sub

Private Sub SendInfo_Close()
bCon = False
CurByte = 0
SendTotal = 0
End Sub

Private Sub SendInfo_Connect()
SendTotal = 0
CurByte = 0
Dim sPack As String
bCon = True
lblStatus = "Status : Requesting Transfer . . ."
sPack = sHeader & sDelim & "F" & sDelim & FileData.FileName & sDelim & FileData.FileSize
SendInfo.SendData sPack
sPack = Empty
End Sub

Private Sub SendInfo_DataArrival(ByVal bytesTotal As Long)
Dim sData As String: sData = Empty
Dim sBuff() As String
SendInfo.GetData sData
Select Case Mid(sData, 12, 1)
Case "F"
    sBuff() = Split(sData, sDelim)
    If sBuff(2) = "Denied" Then
        lblStatus = "Status : User Denied Request."
    ElseIf sBuff(2) = "Accepted" Then
        Dim sRetPack As String
        lblStatus = "Status : Negotiating Transfer . . ."
        SendData.Close
        On Error Resume Next
        SendData.Listen
        sRetPack = sHeader & sDelim & "R"
        SendInfo.SendData sRetPack
    End If
End Select
       
End Sub

Private Sub SendInfo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
CurByte = 0
SendTotal = 0
bCon = False
lblStatus = "Status : Unable to Connect to Recipient."
End Sub

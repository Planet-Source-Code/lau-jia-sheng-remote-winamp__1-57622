VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send a File"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   3480
   End
   Begin MSWinsockLib.Winsock RecInfo 
      Left            =   120
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9010
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send File"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Session Information "
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5415
      Begin MSComDlg.CommonDialog CD 
         Left            =   1080
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblFS 
         Alignment       =   1  'Right Justify
         Caption         =   "File Size (Bytes) :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label lblFN 
         Alignment       =   1  'Right Justify
         Caption         =   "File Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "File to Send :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Recipient's IP Address :"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2040
      End
   End
   Begin MSWinsockLib.Winsock RecData 
      Left            =   120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   4050
   End
   Begin MSWinsockLib.Winsock SendInfo 
      Left            =   720
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9010
   End
   Begin MSWinsockLib.Winsock SendData 
      Left            =   720
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4050
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Provide the IP of the recipient and the file that you want to send."
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4695
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

Private Sub cmdBrowse_Click()
With CD
.DialogTitle = "Select a File to Send"
.Filter = "All Files|*.*"
.Flags = cdlOFNFileMustExist
.ShowOpen
If Len(.FileName) > 0 Then
    If FileLen(.FileName) = 0 Then
        MsgBox "File is empty; please choose another file", vbCritical, "Invalid File"
        Exit Sub
    End If
txtPath.Text = .FileName
FileData.FileName = .FileTitle
FileData.FileSize = FileLen(.FileName)
lblFN.Caption = "File Name : " & FileData.FileName
lblFS.Caption = "File Size (Bytes) : " & FileData.FileSize
End If
End With
End Sub

Private Sub cmdSend_Click()
If Len(txtPath.Text) = 0 Then
MsgBox "Please select a file to send", vbCritical, "File Required"
cmdBrowse_Click
Exit Sub
End If
SendInfo.Close
SendTotal = 0
SaveSetting "File Transfer", "Main", "Host", StrReverse$(txtHost.Text)
bCon = False
SendInfo.Connect txtHost.Text, 9010
lblStatus = "Status : Connecting to Recipient . . ."
End Sub

Private Sub Form_Load()
FileNum = FreeFile
On Error Resume Next
RecInfo.Listen
txtHost.Text = StrReverse$(GetSetting("File Transfer", "Main", "Host", ""))
txtHost_Change
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
Open App.Path & "\Received Files\" & RecFile.FileName For Binary Access Write As #FileNum
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
    iRep = MsgBox(RecInfo.RemoteHostIP & " would like to send you the file " & Chr$(34) & sBuff(2) & Chr$(34) & " (" & sBuff(3) & " Bytes). Accept ?", vbQuestion + vbYesNo, "File Transfer Request")
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

Private Sub SendData_ConnectionRequest(ByVal requestID As Long)
SendData.Close
SendData.Accept requestID
DoEvents
Call SendFile(txtPath.Text)
lblStatus = "Status : Sending File . . ."
End Sub

Private Sub SendData_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
SendTotal = SendTotal + bytesSent
If SendTotal >= FileData.FileSize Then
lblStatus = "Status : File Sent."
SendData.Close

End If

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

Private Sub tmrSend_Timer()
cmdSend_Click
tmrSend.Enabled = False
End Sub

Private Sub txtHost_Change()
If Len(txtHost.Text) = 0 Then
cmdSend.Enabled = False
Else
cmdSend.Enabled = True
End If
End Sub

Private Sub SendFile(sFilePath As String)
Dim FF As Integer: FF = FreeFile
Dim B As Long
Dim bBuffer() As Byte
Open sFilePath For Binary Access Read As #FF
ReDim bBuffer(1 To PacketSize) As Byte
Do Until (FileData.FileSize - CurByte) < PacketSize
DoEvents
Get #FF, CurByte + 1, bBuffer()
CurByte = CurByte + PacketSize
DoEvents
On Error GoTo Err
SendData.SendData bBuffer
Loop
Dim PrevPackSize As Long
PrevPackSize = FileData.FileSize - CurByte
DoEvents
ReDim bBuffer(1 To PrevPackSize) As Byte
Get #FF, CurByte + 1, bBuffer()
CurByte = CurByte + PrevPackSize
DoEvents
SendData.SendData bBuffer
Close #FF
Exit Sub
Err:
Debug.Print "Send Error : " & Err.Description
Exit Sub
End Sub

Private Sub txtPath_Change()
With CD
.FileName = txtPath
FileData.FileName = .FileTitle
FileData.FileSize = FileLen(.FileName)
lblFN.Caption = "File Name : " & FileData.FileName
lblFS.Caption = "File Size (Bytes) : " & FileData.FileSize
End With
End Sub

VERSION 5.00
Begin VB.Form frmfindfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for files in server"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmfindfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   315
      Left            =   9480
      TabIndex        =   4
      Top             =   900
      Width           =   855
   End
   Begin VB.CheckBox chksubfolders 
      Caption         =   "Check sub-folders / Do recursive file search"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   3555
   End
   Begin VB.ListBox lstfoundfiles 
      Height          =   2205
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   10275
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Search"
      Height          =   675
      Left            =   9480
      TabIndex        =   1
      Top             =   160
      Width           =   855
   End
   Begin VB.TextBox txtspec 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "*.mp3"
      Top             =   540
      Width           =   7815
   End
   Begin VB.TextBox txtpath 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "C:\My Music\"
      Top             =   180
      Width           =   7815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Look for:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Search in path:"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmfindfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdfind_Click()
If txtpath = "" Then
Else
frmclient.SendData "FIND" & txtpath & "|" & txtspec & "|" & chksubfolders.Value
End If
End Sub

Private Sub Command1_Click()
lstfoundfiles.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Visible = False
End If
End Sub

Public Function FillinFields(Filepath As String)
txtpath = Filepath
End Function

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub txtpath_GotFocus()
cmdFind.Default = True
End Sub

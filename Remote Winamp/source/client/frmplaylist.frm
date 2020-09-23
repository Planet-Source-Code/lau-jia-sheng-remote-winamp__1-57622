VERSION 5.00
Begin VB.Form frmplaylist 
   Caption         =   "Remote Winamp - View Server Playlist"
   ClientHeight    =   3405
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmplaylist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   7740
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstshortpls 
      Height          =   3180
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7515
   End
   Begin VB.ListBox lstpls 
      Height          =   3180
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   7515
   End
   Begin VB.Menu mnupopupmenu 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnushowfullpath 
         Caption         =   "Show fullpath"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmplaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And vbRightButton Then PopupMenu mnupopupmenu
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
lstpls.Top = lstpls.Left
lstpls.Height = Me.ScaleHeight - (lstpls.Top * 2)
lstpls.Width = Me.ScaleWidth - (lstpls.Left * 2)
lstshortpls.Move lstpls.Left, lstpls.Top, lstpls.Width, lstpls.Height
End Sub

Private Sub lstshortpls_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyExecute Or KeyCode = vbKeyReturn Then lstshortpls_DblClick
End Sub

Private Sub lstshortpls_DblClick()
If lstshortpls.ListIndex <> -1 Then
    frmclient.SendData "GOTO" & lstshortpls.ListIndex
End If
End Sub

Private Sub mnushowfullpath_Click()
    mnushowfullpath.Checked = Not mnushowfullpath.Checked
    lstpls.Visible = mnushowfullpath.Checked
    lstshortpls.Visible = Not mnushowfullpath.Checked
End Sub
Private Sub Form_Initialize()
    InitCommonControls
End Sub

